/**
 * sync-orders-incremental
 *
 * Busca pedidos atualizados nas últimas 48h na Pagar.me (via updated_at)
 * e faz upsert idempotente em orders + order_items.
 *
 * Limites por execução: MAX_PAGES × PAGE_SIZE registros.
 * Designed para terminar em <90s — seguro no timeout de 150s do Supabase.
 *
 * Secrets necessários (Supabase Dashboard → Settings → Edge Functions):
 *   PAGARME_SECRET_KEY  — sk_live_xxxxx
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

// ── Constantes ────────────────────────────────────────────────
const PAGARME_BASE = 'https://api.pagar.me/core/v5'
const PAGE_SIZE    = 100    // máximo suportado pela Pagar.me v5
const MAX_PAGES    = 10     // 1.000 pedidos máx por run (48h com ~20/hora)
const LOOKBACK_H   = 48
const SLEEP_MS     = 350
const BATCH_SIZE   = 100

// ── Classificação de itens (portada do Apps Script) ───────────

function normalizeText(s: unknown): string {
  return String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
}

function parseAddonKeys(code: string): { addon_keys: string[]; is_combo: boolean } {
  const rest = code.replace(/^addon_/, '')

  if (rest.includes('+')) {
    // Combo: chaves separadas por '+', sem sufixo aleatório no final
    const addon_keys = rest
      .split('+')
      .map(s => s.trim().replace(/[^a-z0-9_]/g, ''))
      .filter(Boolean)
    return { addon_keys, is_combo: addon_keys.length > 1 }
  } else {
    // Single: formato KEY_SUFIXO (ex: bin_estadual_DRE6G27) — remove último segmento após '_'
    const lastIdx = rest.lastIndexOf('_')
    const key = (lastIdx >= 0 ? rest.substring(0, lastIdx) : rest)
      .replace(/[^a-z0-9_]/g, '')
    return { addon_keys: key ? [key] : [], is_combo: false }
  }
}

interface ItemClass {
  item_type:   string
  addon_key:   string
  addon_keys:  string[]
  addon_count: number
}

function classifyItem(itemCode: unknown, itemDesc: unknown): ItemClass {
  const code  = String(itemCode ?? '').trim().toLowerCase()
  const descN = normalizeText(itemDesc)

  // Códigos que começam com "addon_" — fonte principal de upsells
  if (code.startsWith('addon_')) {
    const parsed    = parseAddonKeys(code)
    const item_type = parsed.is_combo ? 'BUNDLE' : 'ADDON'
    const addon_key = parsed.is_combo
      ? 'combo:' + parsed.addon_keys.join('+')
      : (parsed.addon_keys[0] ?? '')
    return { item_type, addon_key, addon_keys: parsed.addon_keys, addon_count: parsed.addon_keys.length }
  }

  // Blindagem completa — bundle fixo
  if (descN.includes('blindagem completa')) {
    const addon_keys = [
      'bin_estadual', 'bin_federal', 'gravame', 'historico_leilao',
      'indicio_sinistro', 'dados_proprietario_atual',
    ]
    return {
      item_type:   'BUNDLE',
      addon_key:   'combo:' + addon_keys.join('+'),
      addon_keys,
      addon_count: addon_keys.length,
    }
  }

  // Dados do proprietário
  if (descN.includes('dados do proprietario') || descN.includes('proprietario atual')) {
    return { item_type: 'ADDON', addon_key: 'dados_proprietario_atual',
             addon_keys: ['dados_proprietario_atual'], addon_count: 1 }
  }

  // Item base — não é upsell
  return { item_type: 'OUTRO', addon_key: '', addon_keys: [], addon_count: 0 }
}

// ── Helpers HTTP ──────────────────────────────────────────────

function buildUrl(endpoint: string, params: Record<string, unknown>): string {
  const base = endpoint.startsWith('http') ? endpoint : PAGARME_BASE + endpoint
  const qs = Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== null && String(v) !== '')
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`)
    .join('&')
  return qs ? `${base}?${qs}` : base
}

async function fetchJson(url: string, auth: string): Promise<Record<string, unknown>> {
  const resp = await fetch(url, {
    headers: { Authorization: auth, 'Content-Type': 'application/json' },
  })
  if (!resp.ok) {
    const body = await resp.text()
    throw new Error(`HTTP ${resp.status} → ${url}\n${body}`)
  }
  return resp.json()
}

async function fetchAllPaged(
  auth: string,
  params: Record<string, unknown>,
): Promise<Record<string, unknown>[]> {
  const base = { ...params, size: PAGE_SIZE }
  const out:  Record<string, unknown>[] = []

  for (let page = 1; page <= MAX_PAGES; page++) {
    if (page > 1) await new Promise(r => setTimeout(r, SLEEP_MS))
    const json = await fetchJson(buildUrl('/orders', { ...base, page }), auth)
    const data = Array.isArray(json.data) ? json.data as Record<string, unknown>[] : []
    out.push(...data)
    const paging = json.paging as Record<string, unknown> | undefined
    if (!paging?.next) break
  }

  return out
}

// ── Handler ───────────────────────────────────────────────────

Deno.serve(async () => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const pagarmeKey = Deno.env.get('PAGARME_SECRET_KEY')
  if (!pagarmeKey) {
    return json({ ok: false, error: 'PAGARME_SECRET_KEY não configurada' }, 500)
  }

  const auth        = 'Basic ' + btoa(pagarmeKey + ':')
  const windowEnd   = new Date()
  const windowStart = new Date(windowEnd.getTime() - LOOKBACK_H * 3_600_000)

  // Registra início do run
  const { data: run } = await supabase
    .from('sync_runs')
    .insert({
      integration:  'orders-incremental',
      status:       'running',
      window_start: windowStart.toISOString(),
      window_end:   windowEnd.toISOString(),
    })
    .select('id')
    .single()

  const runId = run?.id
  let recordsProcessed = 0
  let pagesFetched     = 0

  try {
    const orders = await fetchAllPaged(auth, {
      updated_since: windowStart.toISOString(),
      updated_until: windowEnd.toISOString(),
    })

    pagesFetched = Math.ceil(orders.length / PAGE_SIZE)

    for (let i = 0; i < orders.length; i += BATCH_SIZE) {
      const batch = orders.slice(i, i + BATCH_SIZE)

      // ── orders ─────────────────────────────────────────────
      const orderRows = batch.map(o => {
        const c       = o.customer as Record<string, unknown> | undefined
        const charges = Array.isArray(o.charges) ? o.charges as Record<string, unknown>[] : []
        const paidAt  = charges.map(ch => ch.paid_at).find(v => v != null) ?? null
        return {
          provider:           'pagarme',
          provider_order_id:  String(o.id ?? ''),
          order_code:         (o.code as string)     ?? null,
          status:             String(o.status ?? ''),
          created_at:         o.created_at,
          updated_at:         o.updated_at           ?? null,
          paid_at:            paidAt                 ?? null,
          customer_id:        String(c?.id ?? o.customer_id ?? '') || null,
          customer_email:     String(c?.email ?? '')  || null,
          customer_name:      String(c?.name  ?? '')  || null,
          amount:             Number(o.amount  ?? 0),   // centavos — sem dividir
          currency:           String(o.currency ?? 'BRL'),
          raw_json:           o,
          locally_updated_at: new Date().toISOString(),
        }
      })

      const { error: ordErr } = await supabase
        .from('orders')
        .upsert(orderRows, { onConflict: 'provider,provider_order_id' })
      if (ordErr) throw new Error('orders upsert: ' + ordErr.message)

      // ── order_items ────────────────────────────────────────
      const itemRows = batch.flatMap(o => {
        const items = Array.isArray(o.items) ? o.items as Record<string, unknown>[] : []
        return items.map(it => {
          const qty        = Number(it.quantity ?? 0)
          const unitCents  = Number(it.amount   ?? 0)  // campo correto é "amount"
          const totalCents = unitCents * qty
          const cls        = classifyItem(it.code, it.description ?? it.name)
          return {
            provider:           'pagarme',
            provider_order_id:  String(o.id  ?? ''),
            provider_item_id:   String(it.id ?? ''),
            item_code:          String(it.code ?? '')               || null,
            item_description:   String(it.description ?? it.name ?? '') || null,
            item_type:          cls.item_type,
            addon_key:          cls.addon_key  || null,
            addon_keys:         cls.addon_keys,
            qty,
            unit_amount:        unitCents,   // centavos
            total_amount:       totalCents,  // centavos
            raw_json:           it,
          }
        })
      })

      if (itemRows.length > 0) {
        const { error: itmErr } = await supabase
          .from('order_items')
          .upsert(itemRows, { onConflict: 'provider,provider_order_id,provider_item_id' })
        if (itmErr) throw new Error('order_items upsert: ' + itmErr.message)
      }

      recordsProcessed += batch.length
    }

    await supabase.from('sync_runs').update({
      status:            'completed',
      records_processed: recordsProcessed,
      pages_fetched:     pagesFetched,
      finished_at:       new Date().toISOString(),
    }).eq('id', runId)

    return json({ ok: true, records: recordsProcessed, pages: pagesFetched })

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    await supabase.from('sync_runs').update({
      status: 'failed', error_message: msg, finished_at: new Date().toISOString(),
    }).eq('id', runId)
    await supabase.from('sync_errors').insert({
      integration: 'orders-incremental', sync_run_id: runId,
      error_type: 'runtime_error', error_message: msg,
    })
    return json({ ok: false, error: msg }, 500)
  }
})

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  })
}

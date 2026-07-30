/**
 * sync-orders-backfill
 *
 * Processa o histórico de pedidos em janelas de 7 dias (created_at).
 * Lê o cursor em sync_cursors, processa 1 semana, avança o cursor.
 * Para quando o cursor chegar na data atual.
 *
 * Nota: created_since não é suportado pela Pagar.me v5.
 * Usamos apenas created_until + filtro client-side pelo início da janela.
 *
 * Limites: MAX_PAGES × PAGE_SIZE = 1500 pedidos/run (~300/dia × 7 dias).
 * Se a semana tiver mais de 1500 pedidos, status fica 'partial'.
 *
 * Invoke manualmente ou agende com pg_cron.
 *
 * Secrets necessários:
 *   PAGARME_SECRET_KEY — sk_live_xxxxx
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

// ── Constantes ────────────────────────────────────────────────
const PAGARME_BASE  = 'https://api.pagar.me/core/v5'
const PAGE_SIZE     = 100    // ignorado pela Pagar.me v5 (fixo em 30/página)
const MAX_PAGES     = 100    // 3.000 pedidos por run
const SLEEP_MS      = 150    // reduzido para backfill mais rápido
const WINDOW_DAYS   = 3      // janelas de 3 dias — mais granular para períodos densos
const BATCH_SIZE    = 100
const INTEGRATION   = 'orders-backfill'

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

  if (code.startsWith('addon_')) {
    const parsed    = parseAddonKeys(code)
    const item_type = parsed.is_combo ? 'BUNDLE' : 'ADDON'
    const addon_key = parsed.is_combo
      ? 'combo:' + parsed.addon_keys.join('+')
      : (parsed.addon_keys[0] ?? '')
    return { item_type, addon_key, addon_keys: parsed.addon_keys, addon_count: parsed.addon_keys.length }
  }

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

  if (descN.includes('dados do proprietario') || descN.includes('proprietario atual')) {
    return {
      item_type:  'ADDON',
      addon_key:  'dados_proprietario_atual',
      addon_keys: ['dados_proprietario_atual'],
      addon_count: 1,
    }
  }

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

// ── Helpers de data ───────────────────────────────────────────

function addDays(dateStr: string, n: number): string {
  const d = new Date(dateStr + 'T00:00:00Z')
  d.setUTCDate(d.getUTCDate() + n)
  return d.toISOString().slice(0, 10)
}

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function isPastToday(dateStr: string): boolean {
  return dateStr >= todayISO()
}

// ── Handler ───────────────────────────────────────────────────

Deno.serve(async (req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const pagarmeKey = Deno.env.get('PAGARME_SECRET_KEY')
  if (!pagarmeKey) {
    return json({ ok: false, error: 'PAGARME_SECRET_KEY não configurada' }, 500)
  }

  // ── Aceita override de datas no body ─────────────────────────
  // Uso: { "force_start": "2026-02-01", "force_end": "2026-03-01" }
  let forceStart: string | null = null
  let forceEnd:   string | null = null
  try {
    const body = await req.json().catch(() => null)
    if (body && typeof body === 'object') {
      forceStart = body.force_start ?? null
      forceEnd   = body.force_end   ?? null
    }
  } catch { /* body vazio ou null — ok */ }

  // ── Lê cursor atual ──────────────────────────────────────────
  const { data: cursorRow, error: cursorErr } = await supabase
    .from('sync_cursors')
    .select('cursor_value')
    .eq('integration', INTEGRATION)
    .single()

  if (cursorErr || !cursorRow) {
    return json({ ok: false, error: 'Cursor não encontrado para ' + INTEGRATION }, 500)
  }

  const windowStart = forceStart ?? cursorRow.cursor_value

  // Debug: retorna estado atual se ?debug=1
  const url = new URL(req.url)
  if (url.searchParams.get('debug') === '1') {
    return json({
      cursor_in_db: cursorRow.cursor_value,
      window_start: windowStart,
      today:        todayISO(),
      is_done:      isPastToday(windowStart),
      force_start:  forceStart,
    })
  }

  // Para quando chegar em hoje (ignorado se force_start definido)
  if (!forceStart && isPastToday(windowStart)) {
    return json({ ok: true, done: true, message: 'Backfill completo — cursor chegou em hoje.' })
  }

  const windowEnd = forceEnd ?? addDays(windowStart, WINDOW_DAYS)

  // ── Registra início do run ───────────────────────────────────
  const { data: run } = await supabase
    .from('sync_runs')
    .insert({
      integration:  INTEGRATION,
      status:       'running',
      window_start: new Date(windowStart + 'T00:00:00Z').toISOString(),
      window_end:   new Date(windowEnd   + 'T00:00:00Z').toISOString(),
    })
    .select('id')
    .single()

  const runId = run?.id
  const auth  = 'Basic ' + btoa(pagarmeKey + ':')
  let recordsProcessed = 0
  let pagesFetched     = 0
  let hitLimit         = false

  try {
    // ── Busca pedidos com created_until = fim da janela ──────────
    // A API retorna pedidos mais recentes primeiro, dentro de created_until.
    // Filtramos client-side pelo início da janela e paramos quando cruzar windowStart.
    const orders: Record<string, unknown>[] = []
    const windowStartISO = new Date(windowStart + 'T00:00:00Z').toISOString()
    const windowEndISO   = new Date(windowEnd   + 'T00:00:00Z').toISOString()

    for (let page = 1; page <= MAX_PAGES; page++) {
      if (page > 1) await new Promise(r => setTimeout(r, SLEEP_MS))

      const url = buildUrl('/orders', { page, created_until: windowEndISO })
      const json_data = await fetchJson(url, auth)
      const data = Array.isArray(json_data.data)
        ? json_data.data as Record<string, unknown>[]
        : []

      if (data.length === 0) break
      pagesFetched++

      const inWindow = data.filter(o => {
        const ca = String(o.created_at ?? '')
        return ca >= windowStartISO && ca < windowEndISO
      })
      orders.push(...inWindow)

      // Para quando o pedido mais antigo da página cruzar o início da janela
      const oldest = data[data.length - 1]
      if (String(oldest?.created_at ?? '') < windowStartISO) break

      const paging = json_data.paging as Record<string, unknown> | undefined
      if (!paging?.next) break

      if (page === MAX_PAGES) hitLimit = true
    }

    // ── Upsert em lotes ──────────────────────────────────────
    for (let i = 0; i < orders.length; i += BATCH_SIZE) {
      const batch = orders.slice(i, i + BATCH_SIZE)

      const orderRows = batch.map(o => {
        const c = o.customer as Record<string, unknown> | undefined
        return {
          provider:           'pagarme',
          provider_order_id:  String(o.id ?? ''),
          order_code:         (o.code as string) ?? null,
          status:             String(o.status ?? ''),
          created_at:         o.created_at,
          updated_at:         o.updated_at ?? null,
          customer_id:        String(c?.id ?? o.customer_id ?? '') || null,
          customer_email:     String(c?.email ?? '') || null,
          customer_name:      String(c?.name  ?? '') || null,
          amount:             Number(o.amount  ?? 0),
          currency:           String(o.currency ?? 'BRL'),
          raw_json:           o,
          locally_updated_at: new Date().toISOString(),
        }
      })

      const { error: ordErr } = await supabase
        .from('orders')
        .upsert(orderRows, { onConflict: 'provider,provider_order_id' })
      if (ordErr) throw new Error('orders upsert: ' + ordErr.message)

      const itemRows = batch.flatMap(o => {
        const items = Array.isArray(o.items) ? o.items as Record<string, unknown>[] : []
        return items.map(it => {
          const qty        = Number(it.quantity ?? 0)
          const unitCents  = Number(it.amount   ?? 0)
          const totalCents = unitCents * qty
          const cls        = classifyItem(it.code, it.description ?? it.name)
          return {
            provider:           'pagarme',
            provider_order_id:  String(o.id  ?? ''),
            provider_item_id:   String(it.id ?? ''),
            item_code:          String(it.code ?? '')                    || null,
            item_description:   String(it.description ?? it.name ?? '') || null,
            item_type:          cls.item_type,
            addon_key:          cls.addon_key || null,
            addon_keys:         cls.addon_keys,
            qty,
            unit_amount:        unitCents,
            total_amount:       totalCents,
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

    // ── Avança cursor (só se não forçado e não atingiu limite) ──
    const finalStatus = hitLimit ? 'partial' : 'completed'

    if (!forceStart && !hitLimit) {
      await supabase
        .from('sync_cursors')
        .update({ cursor_value: windowEnd, last_updated_at: new Date().toISOString() })
        .eq('integration', INTEGRATION)
    }

    await supabase.from('sync_runs').update({
      status:            finalStatus,
      records_processed: recordsProcessed,
      pages_fetched:     pagesFetched,
      finished_at:       new Date().toISOString(),
    }).eq('id', runId)

    return json({
      ok:      true,
      status:  finalStatus,
      window:  `${windowStart} → ${windowEnd}`,
      records: recordsProcessed,
      pages:   pagesFetched,
      next:    hitLimit ? windowStart : windowEnd,  // próximo cursor
    })

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    await supabase.from('sync_runs').update({
      status: 'failed', error_message: msg, finished_at: new Date().toISOString(),
    }).eq('id', runId)
    await supabase.from('sync_errors').insert({
      integration: INTEGRATION, sync_run_id: runId,
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

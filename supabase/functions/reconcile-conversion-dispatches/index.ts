/**
 * reconcile-conversion-dispatches
 *
 * Job diário de reconciliação (P1.5) — cobre falha/timeout do Database
 * Webhook que dispara `dispatch-conversions`. Lista os checkouts pagos nas
 * últimas 48h no Supabase do SITE (RPC list_paid_checkout_ids) e compara com
 * o que já existe em `conversion_dispatches` no DASHBOARD. Para cada checkout
 * pago sem linha correspondente, roda o MESMO fluxo do dispatch-conversions
 * (RPC get_checkout_conversion_data → dispatchAll → upsert). Limite de
 * segurança: máx. 50 backfills por execução.
 *
 * Secrets necessários (mesmos já usados por dispatch-conversions):
 *   BUREAU_SUPABASE_URL / BUREAU_SUPABASE_KEY
 *   GA4_MEASUREMENT_ID / GA4_API_SECRET
 *   GADS_* / META_PIXEL_ID / META_ACCESS_TOKEN
 *
 * Deploy: supabase functions deploy reconcile-conversion-dispatches --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:   cron-job.org → POST .../functions/v1/reconcile-conversion-dispatches, 1x/dia (08:00 UTC)
 *         Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { dispatchAll, sha256Hex, type CheckoutConversionData } from '../_shared/conversion-dispatch.ts'

const LOOKBACK_HOURS = 48
const MAX_BACKFILL = 50

Deno.serve(async (_req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const bureauUrl = Deno.env.get('BUREAU_SUPABASE_URL')
  const bureauKey = Deno.env.get('BUREAU_SUPABASE_KEY')
  if (!bureauUrl || !bureauKey) {
    return json({ ok: false, error: 'BUREAU_SUPABASE_URL/BUREAU_SUPABASE_KEY não configurados' }, 500)
  }

  const since = new Date(Date.now() - LOOKBACK_HOURS * 3600_000).toISOString()

  try {
    // 1. Lista checkouts pagos nas últimas 48h no site.
    const rpcResp = await fetch(`${bureauUrl}/rest/v1/rpc/list_paid_checkout_ids`, {
      method: 'POST',
      headers: { apikey: bureauKey, Authorization: `Bearer ${bureauKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ p_since: since }),
    })
    if (!rpcResp.ok) throw new Error(`RPC list_paid_checkout_ids HTTP ${rpcResp.status}: ${await rpcResp.text()}`)
    const paidRows = await rpcResp.json() as { checkout_id: string }[]
    const paidIds = paidRows.map(r => r.checkout_id)
    if (!paidIds.length) return json({ ok: true, backfilled: 0 })

    // 2. Busca quais já existem em conversion_dispatches no mesmo período.
    const { data: existingRows, error: existingErr } = await supabase
      .from('conversion_dispatches')
      .select('checkout_id')
      .gte('paid_at', since)
    if (existingErr) throw new Error(`conversion_dispatches select: ${existingErr.message}`)
    const existingIds = new Set((existingRows ?? []).map(r => r.checkout_id))

    const missingIds = paidIds.filter(id => !existingIds.has(id)).slice(0, MAX_BACKFILL)
    if (!missingIds.length) return json({ ok: true, backfilled: 0 })

    // 3. Backfill: mesmo fluxo do dispatch-conversions (RPC → dispatchAll → upsert).
    let backfilled = 0
    for (const checkoutId of missingIds) {
      try {
        const dataResp = await fetch(`${bureauUrl}/rest/v1/rpc/get_checkout_conversion_data`, {
          method: 'POST',
          headers: { apikey: bureauKey, Authorization: `Bearer ${bureauKey}`, 'Content-Type': 'application/json' },
          body: JSON.stringify({ p_checkout_id: checkoutId }),
        })
        if (!dataResp.ok) throw new Error(`RPC get_checkout_conversion_data HTTP ${dataResp.status}: ${await dataResp.text()}`)
        const rows = await dataResp.json() as CheckoutConversionData[]
        if (!rows.length) throw new Error(`checkout ${checkoutId} não encontrado via RPC`)
        const data = rows[0]

        const result = await dispatchAll(data)

        // P2.3 — mesmo padrão do dispatch-conversions: grava só o hash.
        const emailSha256 = data.email ? await sha256Hex(data.email) : null
        const phoneSha256 = data.phone ? await sha256Hex(data.phone.replace(/\D/g, '')) : null

        await supabase.from('conversion_dispatches').upsert({
          checkout_id: checkoutId,
          order_nsu:   data.order_nsu,
          event_id:    result.eventId,
          value:       data.valor,
          currency:    'BRL',
          paid_at:     data.paid_at ?? null,
          email_sha256: emailSha256,
          phone_sha256: phoneSha256,
          gclid:             data.gclid ?? null,
          gbraid:            data.gbraid ?? null,
          wbraid:            data.wbraid ?? null,
          fbp:               data.fbp ?? null,
          fbc:               data.fbc ?? null,
          ga_client_id:      data.ga_client_id ?? null,
          event_source_url:  data.event_source_url ?? null,
          client_user_agent: data.client_user_agent ?? null,
          ga4_status:  result.ga4.status,  ga4_error:  result.ga4.error  ?? null,
          ads_status:  result.ads.status,  ads_error:  result.ads.error  ?? null,
          meta_status: result.meta.status, meta_error: result.meta.error ?? null,
          attempts:    1,
          updated_at:  new Date().toISOString(),
        }, { onConflict: 'checkout_id' })

        backfilled++
      } catch (err) {
        // Não interrompe o loop — próxima execução (24h depois) tenta de novo,
        // e o retry-conversion-dispatches cobre falhas parciais de canal.
        console.error(`reconcile: falha no checkout ${checkoutId}:`, err instanceof Error ? err.message : String(err))
      }
    }

    return json({ ok: true, backfilled })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    return json({ ok: false, error: msg }, 500)
  }
})

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { 'Content-Type': 'application/json' } })
}

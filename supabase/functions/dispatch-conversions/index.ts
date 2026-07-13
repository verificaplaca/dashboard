/**
 * dispatch-conversions
 *
 * Recebe o Supabase Database Webhook do projeto do SITE (ozquoloetuzynnyzkado)
 * disparado em UPDATE de checkouts, e envia Purchase pra GA4 MP + Google Ads
 * Enhanced Conversions for Leads + Meta CAPI. Grava status por canal em
 * conversion_dispatches (idempotente por checkout_id) pro job de retry cobrir falhas.
 *
 * Gatilho correto: transição paid_at NULL → preenchido (não status='paid' —
 * confirmado em dados reais que status='paid' aparece às vezes com paid_at nulo).
 *
 * Secrets necessários (Supabase → Project Settings → Edge Functions → Secrets,
 * projeto ftmgmfdqdqxboiktxcoj = dashboard):
 *   BUREAU_SUPABASE_URL / BUREAU_SUPABASE_KEY — já existem (reaproveitados do pipeline de bureau)
 *   GA4_MEASUREMENT_ID / GA4_API_SECRET       — configurados nesta sessão
 *   GADS_CLIENT_ID / GADS_CLIENT_SECRET / GADS_REFRESH_TOKEN / GADS_DEVELOPER_TOKEN / GADS_CUSTOMER_ID — já existem
 *   GADS_PURCHASE_CONVERSION_ACTION_ID — PENDENTE, ver instrucoes-setup-passos-5-7.md
 *   META_PIXEL_ID / META_ACCESS_TOKEN  — PENDENTE, usuário vai habilitar depois
 *
 * Deploy: supabase functions deploy dispatch-conversions --project-ref ftmgmfdqdqxboiktxcoj
 *
 * Modo diagnóstico: GET .../dispatch-conversions?diag=gads_conversion_actions
 *   lista as conversion actions do Google Ads (reaproveitando os secrets GADS_* já
 *   configurados) pra descobrir o conversion_action_id numérico da ação "purchase" —
 *   o Conversion ID/label que o Google Ads mostra na tela (AW-xxx/label) NÃO é o
 *   mesmo ID numérico que a API precisa.
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { dispatchAll, type CheckoutConversionData } from '../_shared/conversion-dispatch.ts'

Deno.serve(async (req) => {
  const url = new URL(req.url)

  // ── Modo diagnóstico: lista conversion actions do Google Ads ─────────────
  if (req.method === 'GET' && url.searchParams.get('diag') === 'gads_conversion_actions') {
    return await diagListConversionActions()
  }

  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const bureauUrl = Deno.env.get('BUREAU_SUPABASE_URL')
  const bureauKey = Deno.env.get('BUREAU_SUPABASE_KEY')
  if (!bureauUrl || !bureauKey) {
    return json({ ok: false, error: 'BUREAU_SUPABASE_URL/BUREAU_SUPABASE_KEY não configurados' }, 500)
  }

  let payload: any
  try {
    payload = await req.json()
  } catch {
    return json({ ok: false, error: 'body inválido' }, 400)
  }

  // Payload do Supabase Database Webhook: { type, table, record, old_record, schema }
  if (payload.table !== 'checkouts' || payload.type !== 'UPDATE') {
    return json({ ok: true, skipped: 'não é UPDATE em checkouts' })
  }
  const oldPaid = payload.old_record?.paid_at
  const newPaid = payload.record?.paid_at
  if (oldPaid != null || newPaid == null) {
    return json({ ok: true, skipped: 'não é transição paid_at NULL → preenchido' })
  }

  const checkoutId = payload.record.id as string

  try {
    // Busca dados do checkout + email/phone via RPC no Supabase do site
    // (site_get_checkout_conversion_data.sql precisa estar deployado lá).
    const rpcResp = await fetch(`${bureauUrl}/rest/v1/rpc/get_checkout_conversion_data`, {
      method: 'POST',
      headers: { apikey: bureauKey, Authorization: `Bearer ${bureauKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ p_checkout_id: checkoutId }),
    })
    if (!rpcResp.ok) throw new Error(`RPC get_checkout_conversion_data HTTP ${rpcResp.status}: ${await rpcResp.text()}`)
    const rows = await rpcResp.json() as CheckoutConversionData[]
    if (!rows.length) throw new Error(`checkout ${checkoutId} não encontrado via RPC`)
    const data = rows[0]

    const result = await dispatchAll(data)

    await supabase.from('conversion_dispatches').upsert({
      checkout_id: checkoutId,
      order_nsu:   data.order_nsu,
      event_id:    result.eventId,
      value:       data.valor,
      currency:    'BRL',
      email:       data.email ?? null,
      phone:       data.phone ?? null,
      ga4_status:  result.ga4.status,  ga4_error:  result.ga4.error  ?? null,
      ads_status:  result.ads.status,  ads_error:  result.ads.error  ?? null,
      meta_status: result.meta.status, meta_error: result.meta.error ?? null,
      attempts:    1,
      updated_at:  new Date().toISOString(),
    }, { onConflict: 'checkout_id' })

    return json({ ok: true, checkout_id: checkoutId, result })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    await supabase.from('conversion_dispatches').upsert({
      checkout_id: checkoutId,
      order_nsu:   payload.record.order_nsu ?? '',
      event_id:    'unknown',
      value:       payload.record.valor ?? 0,
      currency:    'BRL',
      ga4_status: 'failed', ga4_error: msg,
      ads_status: 'failed', ads_error: msg,
      meta_status: 'failed', meta_error: msg,
      attempts: 1,
      updated_at: new Date().toISOString(),
    }, { onConflict: 'checkout_id' })
    return json({ ok: false, error: msg }, 500)
  }
})

async function diagListConversionActions(): Promise<Response> {
  const clientId     = Deno.env.get('GADS_CLIENT_ID')
  const clientSecret = Deno.env.get('GADS_CLIENT_SECRET')
  const refreshToken = Deno.env.get('GADS_REFRESH_TOKEN')
  const devToken      = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId    = Deno.env.get('GADS_CUSTOMER_ID')
  if (!clientId || !clientSecret || !refreshToken || !devToken || !customerId) {
    return json({ ok: false, error: 'secrets GADS_* não configurados' }, 500)
  }
  try {
    const tokenResp = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken, grant_type: 'refresh_token' }),
    })
    if (!tokenResp.ok) throw new Error('OAuth error: ' + await tokenResp.text())
    const { access_token } = await tokenResp.json()

    const query = `SELECT conversion_action.id, conversion_action.name, conversion_action.type, conversion_action.status FROM conversion_action`
    const resp = await fetch(`https://googleads.googleapis.com/v24/customers/${customerId}/googleAds:search`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${access_token}`, 'developer-token': devToken, 'Content-Type': 'application/json' },
      body: JSON.stringify({ query }),
    })
    if (!resp.ok) throw new Error(`Google Ads API HTTP ${resp.status}: ${await resp.text()}`)
    const data = await resp.json()
    return json({ ok: true, conversion_actions: data.results ?? [] })
  } catch (err) {
    return json({ ok: false, error: err instanceof Error ? err.message : String(err) }, 500)
  }
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { 'Content-Type': 'application/json' } })
}

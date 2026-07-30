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
import {
  buildEnhancementAdjustment,
  dispatchAll,
  formatConversionDateTime,
  getGadsAccessToken,
  sha256Hex,
  type CheckoutConversionData,
} from '../_shared/conversion-dispatch.ts'

Deno.serve(async (req) => {
  const url = new URL(req.url)

  // ── Modo diagnóstico: lista conversion actions do Google Ads ─────────────
  if (req.method === 'GET' && url.searchParams.get('diag') === 'gads_conversion_actions') {
    return await diagListConversionActions()
  }

  // ── Modo diagnóstico: quais secrets GADS_* estão configurados ────────────
  // Devolve só booleanos (ou o ID, que não é segredo) — nunca client secret ou
  // refresh token. Responde na hora se o caminho de Click Conversions está
  // ativo: quando GADS_CLICK_CONVERSION_ACTION_ID existe, todo checkout com
  // gclid pula o Enhanced Conversions e sobe SEM nenhum userIdentifier
  // (early return no dispatchGoogleAds, em _shared/conversion-dispatch.ts).
  if (req.method === 'GET' && url.searchParams.get('diag') === 'secrets') {
    const has = (k: string) => Boolean(Deno.env.get(k))
    return json({
      ok: true,
      secrets: {
        GADS_CUSTOMER_ID:                   has('GADS_CUSTOMER_ID'),
        GADS_DEVELOPER_TOKEN:               has('GADS_DEVELOPER_TOKEN'),
        GADS_REFRESH_TOKEN:                 has('GADS_REFRESH_TOKEN'),
        GADS_PURCHASE_CONVERSION_ACTION_ID: Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID') ?? null,
        GADS_CLICK_CONVERSION_ACTION_ID:    Deno.env.get('GADS_CLICK_CONVERSION_ACTION_ID') ?? null,
        GADS_EC_SEND_CONVERSION_DATE_TIME:  Deno.env.get('GADS_EC_SEND_CONVERSION_DATE_TIME') ?? null,
        GA4_MEASUREMENT_ID:                 has('GA4_MEASUREMENT_ID'),
        META_PIXEL_ID:                      has('META_PIXEL_ID'),
      },
      nota: 'GADS_CLICK_CONVERSION_ACTION_ID preenchido => caminho de Click Conversions ativo para checkouts com gclid, e esses NAO levam user identifiers.',
    })
  }

  // ── Modo diagnóstico: diagnósticos de upload offline direto da API do Ads ─
  // Fonte da verdade pros alertas do painel "Needs attention", com contagem por
  // dia (daily_summaries) em vez de interpretação da UI.
  // Doc: developers.google.com/google-ads/api/docs/conversions/upload-summaries
  if (req.method === 'GET' && url.searchParams.get('diag') === 'upload_summary') {
    return await diagUploadSummary(url.searchParams.get('conversion_action_id'))
  }

  // ── Modo diagnóstico: dry run do payload de enhancement (validateOnly) ────
  // Manda o payload REAL (mesma função buildEnhancementAdjustment usada pelo
  // dispatch) com validateOnly=true: o Google valida e NAO grava nada. Rodar
  // antes de ligar GADS_EC_SEND_CONVERSION_DATE_TIME=1 em produção.
  // Uso: ?diag=enhancement_dryrun&order_id=order_xxxxxxxx
  if (req.method === 'GET' && url.searchParams.get('diag') === 'enhancement_dryrun') {
    return await diagEnhancementDryRun(url.searchParams.get('order_id'))
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

    // P2.3 — PII hash at rest: grava só o SHA-256 de email/phone.
    // Telefone em DOIS formatos (migration 015), porque cada canal exige o seu e
    // o retry consome o hash pronto: a RPC do site devolve E.164 com '+'.
    //   _e164 → Google Ads | _digits → Meta CAPI
    const emailSha256       = data.email ? await sha256Hex(data.email) : null
    const phoneSha256E164   = data.phone ? await sha256Hex(data.phone) : null
    const phoneSha256Digits = data.phone ? await sha256Hex(data.phone.replace(/\D/g, '')) : null

    await supabase.from('conversion_dispatches').upsert({
      checkout_id: checkoutId,
      order_nsu:   data.order_nsu,
      event_id:    result.eventId,
      value:       data.valor,
      currency:    'BRL',
      paid_at:     data.paid_at ?? null,
      email_sha256:        emailSha256,
      phone_sha256_e164:   phoneSha256E164,
      phone_sha256_digits: phoneSha256Digits,
      // P1.4 — atribuição capturada no client (P1.1); persistida aqui pro
      // retry não depender de nova chamada de RPC no site.
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

    return json({ ok: true, checkout_id: checkoutId, result })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    await supabase.from('conversion_dispatches').upsert({
      checkout_id: checkoutId,
      order_nsu:   payload.record.order_nsu ?? '',
      event_id:    'unknown',
      value:       payload.record.valor ?? 0,
      currency:    'BRL',
      paid_at:     payload.record.paid_at ?? null,
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

// ── diag=upload_summary ─────────────────────────────────────────────────────
// Puxa o mesmo diagnóstico que o painel "Needs attention" do Google Ads mostra,
// mas via API e com quebra por dia. É o que permite separar "problema vivo" de
// "resíduo de dado antigo ainda dentro da janela do painel".
async function diagUploadSummary(conversionActionIdParam: string | null): Promise<Response> {
  const devToken   = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId = Deno.env.get('GADS_CUSTOMER_ID')
  if (!devToken || !customerId) return json({ ok: false, error: 'secrets GADS_* não configurados' }, 500)

  const actionId = conversionActionIdParam ?? Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID') ?? null

  const R = 'offline_conversion_upload_conversion_action_summary'
  const fields = [
    `${R}.conversion_action_id`,
    `${R}.conversion_action_name`,
    `${R}.client`,
    `${R}.status`,
    `${R}.alerts`,
    `${R}.total_event_count`,
    `${R}.successful_event_count`,
    `${R}.pending_event_count`,
    `${R}.last_upload_date_time`,
    `${R}.daily_summaries`,
    `${R}.job_summaries`,
  ]
  let query = `SELECT ${fields.join(', ')} FROM ${R}`
  if (actionId) query += ` WHERE ${R}.conversion_action_id = ${actionId}`

  // Nível CONTA. É aqui que aparecem success_rate/pending_rate e, principalmente,
  // o campo `alerts` agregado — o resumo por conversion action veio sem alerts
  // nenhum em 30/07, então vale olhar se o nível de conta mostra algo.
  const C = 'offline_conversion_upload_client_summary'
  const clientQuery = `SELECT ${[
    `${C}.client`,
    `${C}.status`,
    `${C}.alerts`,
    `${C}.total_event_count`,
    `${C}.successful_event_count`,
    `${C}.pending_event_count`,
    `${C}.success_rate`,
    `${C}.pending_rate`,
    `${C}.last_upload_date_time`,
    `${C}.daily_summaries`,
  ].join(', ')} FROM ${C}`

  try {
    const accessToken = await getGadsAccessToken()
    const run = async (q: string) => {
      const resp = await fetch(`https://googleads.googleapis.com/v24/customers/${customerId}/googleAds:search`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, 'developer-token': devToken, 'Content-Type': 'application/json' },
        body: JSON.stringify({ query: q }),
      })
      const text = await resp.text()
      let parsed: unknown
      try { parsed = JSON.parse(text) } catch { parsed = text }
      return { http_status: resp.status, data: parsed }
    }
    const [porAcao, porConta] = await Promise.all([run(query), run(clientQuery)])
    return json({
      ok: true,
      agora_utc: new Date().toISOString(),
      por_conversion_action: { query, ...porAcao },
      por_conta: { query: clientQuery, ...porConta },
    })
  } catch (err) {
    return json({ ok: false, query, error: err instanceof Error ? err.message : String(err) }, 500)
  }
}

// ── diag=enhancement_dryrun ─────────────────────────────────────────────────
// Valida 3 formatos de payload contra a API real com validateOnly=true (nada é
// gravado no Google). Responde, com evidência, se dá pra ligar
// GADS_EC_SEND_CONVERSION_DATE_TIME=1 sem quebrar o dispatch em produção.
//   A = payload como estava até 30/07 (sem NENHUM campo de tempo) — baseline
//   B = payload novo: adjustmentDateTime + userAgent
//   C = B + gclidDateTimePair.conversionDateTime (o campo em dúvida na doc)
async function diagEnhancementDryRun(orderId: string | null): Promise<Response> {
  const devToken           = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId         = Deno.env.get('GADS_CUSTOMER_ID')
  const conversionActionId = Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID')
  if (!devToken || !customerId || !conversionActionId) {
    return json({ ok: false, error: 'GADS_CUSTOMER_ID / GADS_DEVELOPER_TOKEN / GADS_PURCHASE_CONVERSION_ACTION_ID não configurados' }, 500)
  }
  if (!orderId) return json({ ok: false, error: 'informe ?order_id=<order_nsu de uma compra real recente>' }, 400)

  const nowIso  = new Date().toISOString()
  const paidIso = new Date(Date.now() - 2 * 3600_000).toISOString()

  try {
    const accessToken = await getGadsAccessToken()
    // Identificador fictício: com validateOnly o Google checa a ESTRUTURA, não
    // se o hash casa com alguém. Não usar PII real num endpoint de diagnóstico.
    const userIdentifiers = [{ hashedEmail: await sha256Hex('dryrun@verificaplaca.com.br') }]

    const fake: CheckoutConversionData = {
      checkout_id: '00000000-0000-0000-0000-000000000000',
      order_nsu: orderId,
      valor: 14.99,
      paid_at: paidIso,
      client_user_agent: 'Mozilla/5.0 (dry-run)',
    }

    const b = buildEnhancementAdjustment(fake, customerId, conversionActionId, userIdentifiers, nowIso)
    delete b.gclidDateTimePair // variante B nunca leva o campo em dúvida
    const c = { ...b, gclidDateTimePair: { conversionDateTime: formatConversionDateTime(paidIso) } }
    const a = {
      conversionAction: `customers/${customerId}/conversionActions/${conversionActionId}`,
      adjustmentType: 'ENHANCEMENT',
      orderId,
      userIdentifiers,
    }

    const variants: Record<string, unknown> = { A_baseline_sem_tempo: a, B_adjustmentDateTime_userAgent: b, C_B_mais_conversionDateTime: c }
    const results: Record<string, unknown> = {}

    for (const [nome, adjustment] of Object.entries(variants)) {
      const resp = await fetch(
        `https://googleads.googleapis.com/v24/customers/${customerId}:uploadConversionAdjustments`,
        {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, 'developer-token': devToken, 'Content-Type': 'application/json' },
          body: JSON.stringify({ conversionAdjustments: [adjustment], partialFailure: true, validateOnly: true }),
        }
      )
      const text = await resp.text()
      let parsed: unknown
      try { parsed = JSON.parse(text) } catch { parsed = text }
      results[nome] = { http_status: resp.status, enviado: adjustment, resposta: parsed }
    }

    return json({
      ok: true,
      aviso: 'validateOnly=true — nada foi gravado no Google Ads.',
      como_ler: 'Variante sem partialFailureError e com http 200 é aceita. Se C passar, pode ligar GADS_EC_SEND_CONVERSION_DATE_TIME=1. Se C falhar, deixar a flag desligada e ficar só com B.',
      results,
    })
  } catch (err) {
    return json({ ok: false, error: err instanceof Error ? err.message : String(err) }, 500)
  }
}

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { 'Content-Type': 'application/json' } })
}

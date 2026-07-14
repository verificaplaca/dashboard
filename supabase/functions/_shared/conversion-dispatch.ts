/**
 * conversion-dispatch.ts
 *
 * Lógica compartilhada de dispatch de conversão de Purchase pra 3 canais:
 * GA4 Measurement Protocol, Google Ads Enhanced Conversions for Leads, Meta CAPI.
 * Usada por dispatch-conversions (webhook) e retry-conversion-dispatches (cron).
 *
 * ⚠️ Google Ads: a chamada de Enhanced Conversions for Leads abaixo (ConversionAdjustmentUploadService.
 * uploadConversionAdjustments, adjustment_type ENHANCEMENT) é a forma que a documentação pública descreve,
 * MAS não foi validada ao vivo nesta sessão — antes do primeiro deploy real, confirmar o formato exato do
 * request contra https://developers.google.com/google-ads/api/docs/conversions/enhanced-conversions/leads
 * (nomes de campo podem ter mudado entre versões da API). Não tratar esse trecho como testado.
 */

export interface CheckoutConversionData {
  checkout_id: string
  order_nsu: string
  placa?: string
  plano?: string
  valor: number
  transaction_id?: string | null
  email?: string | null
  phone?: string | null
  paid_at?: string | null
  // P2.3 — hash pronto (SHA-256), quando o dado vem do banco (retry/reconcile
  // não guardam mais email/phone cru, só o hash). Quando ausente, os
  // dispatchers hasheiam `email`/`phone` na hora (fluxo fresco via RPC do site).
  email_sha256?: string | null
  phone_sha256?: string | null
  // P1.4 — atribuição capturada no client do site (P1.1). Tudo opcional:
  // checkouts anteriores ao deploy do site vêm com esses campos null, e o
  // dispatch precisa continuar funcionando com os fallbacks atuais.
  gclid?: string | null
  gbraid?: string | null
  wbraid?: string | null
  fbp?: string | null
  fbc?: string | null
  ga_client_id?: string | null
  event_source_url?: string | null
  client_user_agent?: string | null
}

export interface DispatchResult {
  ga4:  { status: 'success' | 'failed' | 'skipped'; error?: string }
  ads:  { status: 'success' | 'failed' | 'skipped'; error?: string }
  meta: { status: 'success' | 'failed' | 'skipped'; error?: string }
}

// event_id determinístico — precisa ser IDÊNTICO ao gerado no client (analytics.ts):
//   - fluxo principal (Sucesso.tsx):        uuid: orderNsu           → 'evt_' + order_nsu
//   - fluxo addon (CheckoutAddon/Resultado3): uuid: pixData.checkout_id → 'evt_' + checkout_id
// Distinguido pelo prefixo do order_nsu ('addon_' vs 'order_'), confirmado nos dados reais.
export function computeEventId(row: { checkout_id: string; order_nsu: string }): string {
  return row.order_nsu?.startsWith('addon_')
    ? 'evt_' + row.checkout_id
    : 'evt_' + row.order_nsu
}

// transaction_id/orderId determinístico — precisa ser IDÊNTICO ao que o client envia:
//   - fluxo principal (Sucesso.tsx):          transaction_id = orderNsu
//   - fluxo addon (CheckoutAddon/Resultado3): transaction_id = pixData.checkout_id
// Sem isso, GA4 vê 2 transações (receita de addon duplicada) e o dedup por
// orderId no Google Ads nunca casa em addons.
export function computeTransactionId(row: { checkout_id: string; order_nsu: string }): string {
  return row.order_nsu?.startsWith('addon_') ? row.checkout_id : row.order_nsu
}

// Exportado: dispatch-conversions/reconcile usam pra gravar email_sha256/
// phone_sha256 em conversion_dispatches (P2.3 — não guarda mais PII em claro).
export async function sha256Hex(input: string): Promise<string> {
  const data = new TextEncoder().encode(input.trim().toLowerCase())
  const hashBuffer = await crypto.subtle.digest('SHA-256', data)
  return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join('')
}

// ── GA4 Measurement Protocol ────────────────────────────────────────────────
// Docs: https://developers.google.com/analytics/devguides/collection/protocol/ga4
async function dispatchGA4(data: CheckoutConversionData, eventId: string): Promise<DispatchResult['ga4']> {
  const measurementId = Deno.env.get('GA4_MEASUREMENT_ID')
  const apiSecret      = Deno.env.get('GA4_API_SECRET')
  if (!measurementId || !apiSecret) return { status: 'skipped', error: 'GA4_MEASUREMENT_ID/GA4_API_SECRET não configurados' }

  // GA4 MP exige client_id. P1.4: se o site capturou o _ga real (ga_client_id),
  // o purchase server-side se junta à sessão/atribuição original do usuário.
  // Sem isso (checkouts antigos ou site ainda não deployado com P1.1), cai no
  // fallback synthetic determinístico a partir do checkout_id — o hit é
  // registrado no GA4 mas NÃO se junta à sessão original (limitação conhecida, não é bug).
  const clientId = data.ga_client_id || `srv.${data.checkout_id.replace(/-/g, '').slice(0, 16)}`

  const body = {
    client_id: clientId,
    // Hora real da conversão (paid_at), não a hora do dispatch/retry.
    // GA4 MP aceita eventos retroativos até ~72h — retries dentro da janela
    // do job (5 tentativas x 15min) ficam muito abaixo disso.
    timestamp_micros: data.paid_at ? new Date(data.paid_at).getTime() * 1000 : undefined,
    events: [{
      name: 'purchase',
      params: {
        transaction_id: computeTransactionId(data),
        value: data.valor,
        currency: 'BRL',
        plan_name: data.plano,
        event_id: eventId,
      },
    }],
  }

  try {
    const resp = await fetch(
      `https://www.google-analytics.com/mp/collect?measurement_id=${measurementId}&api_secret=${apiSecret}`,
      { method: 'POST', body: JSON.stringify(body) }
    )
    // GA4 MP retorna 204 mesmo em payload malformado (não valida sincronamente) —
    // só HTTP status != 2xx indica erro de fato (ex: measurement_id/api_secret inválidos).
    if (!resp.ok) return { status: 'failed', error: `GA4 MP HTTP ${resp.status}: ${await resp.text()}` }
    return { status: 'success' }
  } catch (err) {
    return { status: 'failed', error: err instanceof Error ? err.message : String(err) }
  }
}

// ── Google Ads — Enhanced Conversions for Leads ─────────────────────────────
async function getGadsAccessToken(): Promise<string> {
  const clientId     = Deno.env.get('GADS_CLIENT_ID')!
  const clientSecret = Deno.env.get('GADS_CLIENT_SECRET')!
  const refreshToken = Deno.env.get('GADS_REFRESH_TOKEN')!
  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken, grant_type: 'refresh_token' }),
  })
  if (!resp.ok) throw new Error('OAuth token error: ' + await resp.text())
  return (await resp.json() as { access_token: string }).access_token
}

// Formata paid_at (ISO) no formato exigido pela API: 'yyyy-MM-dd HH:mm:ss+00:00'.
// Referência (exemplo oficial da doc): '2021-01-01 12:32:45-08:00'. Usamos sempre
// +00:00 porque paid_at é armazenado/lido como timestamptz (UTC) no Postgres.
function formatConversionDateTime(paidAt: string): string {
  const d = new Date(paidAt)
  const pad = (n: number) => String(n).padStart(2, '0')
  const y = d.getUTCFullYear()
  const mo = pad(d.getUTCMonth() + 1)
  const day = pad(d.getUTCDate())
  const h = pad(d.getUTCHours())
  const mi = pad(d.getUTCMinutes())
  const s = pad(d.getUTCSeconds())
  return `${y}-${mo}-${day} ${h}:${mi}:${s}+00:00`
}

// ── Google Ads — Click Conversions (com gclid/gbraid/wbraid) ───────────────
// Doc: https://developers.google.com/google-ads/api/docs/conversions/upload-clicks
// Confirmado via exemplo oficial (google-ads-python, examples/remarketing/upload_offline_conversion.py,
// v24): campos do objeto ClickConversion em REST/JSON (camelCase) são
// conversionAction, gclid | gbraid | wbraid (exatamente um dos três — não
// combinar), conversionValue, conversionDateTime, currencyCode, orderId.
// partialFailure vai no nível do request (mesmo padrão do uploadConversionAdjustments
// já usado abaixo). consent é opcional (usado em jurisdições com Consent Mode) —
// NÃO VALIDADO NA DOC se é obrigatório para esta conta; omitido por padrão.
async function dispatchGoogleAdsClickConversion(
  data: CheckoutConversionData,
  customerId: string,
  devToken: string,
  conversionActionId: string,
): Promise<DispatchResult['ads']> {
  if (!data.paid_at) return { status: 'failed', error: 'paid_at ausente — obrigatório pra conversionDateTime' }

  try {
    const accessToken = await getGadsAccessToken()

    const clickConversion: Record<string, unknown> = {
      conversionAction: `customers/${customerId}/conversionActions/${conversionActionId}`,
      conversionDateTime: formatConversionDateTime(data.paid_at),
      conversionValue: data.valor,
      currencyCode: 'BRL',
      orderId: computeTransactionId(data),
    }
    // Exatamente um identificador de clique — prioridade gclid > gbraid > wbraid
    // (mesma ordem do exemplo oficial; gclid nunca coexiste com os outros dois).
    if (data.gclid) clickConversion.gclid = data.gclid
    else if (data.gbraid) clickConversion.gbraid = data.gbraid
    else if (data.wbraid) clickConversion.wbraid = data.wbraid

    const resp = await fetch(
      `https://googleads.googleapis.com/v24/customers/${customerId}:uploadClickConversions`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'developer-token': devToken,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          conversions: [clickConversion],
          partialFailure: true,
        }),
      }
    )
    if (!resp.ok) return { status: 'failed', error: `Google Ads API (uploadClickConversions) HTTP ${resp.status}: ${await resp.text()}` }
    const json = await resp.json()
    if (json.partialFailureError) return { status: 'failed', error: JSON.stringify(json.partialFailureError) }
    return { status: 'success' }
  } catch (err) {
    return { status: 'failed', error: err instanceof Error ? err.message : String(err) }
  }
}

async function dispatchGoogleAds(data: CheckoutConversionData): Promise<DispatchResult['ads']> {
  const customerId       = Deno.env.get('GADS_CUSTOMER_ID')
  const devToken          = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const conversionActionId = Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID') // ID numérico — ainda não confirmado, ver doc de setup
  if (!customerId || !devToken || !conversionActionId) {
    return { status: 'skipped', error: 'GADS_PURCHASE_CONVERSION_ACTION_ID não configurado (ver instrucoes-setup-passos-5-7.md)' }
  }

  // P1.4: se o site capturou algum identificador de clique, usar Click
  // Conversions (uploadClickConversions). ⚠️ Esse endpoint SÓ aceita
  // conversion actions do tipo "Importar > conversões de cliques" — a action
  // da tag do site (webpage, GADS_PURCHASE_CONVERSION_ACTION_ID) retorna
  // INVALID_CONVERSION_ACTION_TYPE (visto em produção 14/07/2026). Por isso
  // o caminho só ativa quando GADS_CLICK_CONVERSION_ACTION_ID (action de
  // importação dedicada, criada no Google Ads como ação secundária) estiver
  // configurada. Sem ela, cai no Enhanced Conversions abaixo.
  const clickConversionActionId = Deno.env.get('GADS_CLICK_CONVERSION_ACTION_ID')
  if (clickConversionActionId && (data.gclid || data.gbraid || data.wbraid)) {
    return dispatchGoogleAdsClickConversion(data, customerId, devToken, clickConversionActionId)
  }

  // Fallback: sem GADS_CLICK_CONVERSION_ACTION_ID configurada ou sem id de
  // clique — caminho de Enhanced Conversions (validado em produção).
  if (!data.email && !data.phone && !data.email_sha256 && !data.phone_sha256) {
    return { status: 'skipped', error: 'checkout sem email/phone (auth.users) — Enhanced Conversions for Leads exige ao menos um identificador' }
  }

  try {
    const accessToken = await getGadsAccessToken()
    const userIdentifiers: Record<string, unknown>[] = []
    // P2.3: usa o hash já gravado (retry/reconcile) quando disponível; senão
    // hasheia o dado cru vindo da RPC (fluxo fresco do dispatch-conversions).
    // Normalização de phone preservada igual ao comportamento anterior (sem
    // strip de não-dígitos) quando o hash é calculado aqui.
    if (data.email_sha256) userIdentifiers.push({ hashedEmail: data.email_sha256 })
    else if (data.email) userIdentifiers.push({ hashedEmail: await sha256Hex(data.email) })
    if (data.phone_sha256) userIdentifiers.push({ hashedPhoneNumber: data.phone_sha256 })
    else if (data.phone) userIdentifiers.push({ hashedPhoneNumber: await sha256Hex(data.phone) })

    // ⚠️ Ver aviso no topo do arquivo — formato não validado ao vivo.
    const resp = await fetch(
      `https://googleads.googleapis.com/v24/customers/${customerId}:uploadConversionAdjustments`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'developer-token': devToken,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          conversionAdjustments: [{
            conversionAction: `customers/${customerId}/conversionActions/${conversionActionId}`,
            adjustmentType: 'ENHANCEMENT',
            orderId: computeTransactionId(data),
            userIdentifiers,
          }],
          partialFailure: true,
        }),
      }
    )
    if (!resp.ok) return { status: 'failed', error: `Google Ads API HTTP ${resp.status}: ${await resp.text()}` }
    const json = await resp.json()
    if (json.partialFailureError) return { status: 'failed', error: JSON.stringify(json.partialFailureError) }
    return { status: 'success' }
  } catch (err) {
    return { status: 'failed', error: err instanceof Error ? err.message : String(err) }
  }
}

// ── Meta Conversions API ────────────────────────────────────────────────────
// Docs: https://developers.facebook.com/docs/marketing-api/conversions-api
async function dispatchMeta(data: CheckoutConversionData, eventId: string): Promise<DispatchResult['meta']> {
  const pixelId     = Deno.env.get('META_PIXEL_ID')
  const accessToken = Deno.env.get('META_ACCESS_TOKEN')
  if (!pixelId || !accessToken) return { status: 'skipped', error: 'META_PIXEL_ID/META_ACCESS_TOKEN não configurados ainda (pendente de habilitar Conversions API)' }

  const userData: Record<string, unknown> = {}
  // P2.3: usa o hash já gravado (retry/reconcile) quando disponível; senão
  // hasheia o dado cru vindo da RPC (fluxo fresco do dispatch-conversions),
  // preservando a normalização anterior (phone sem não-dígitos antes do hash).
  if (data.email_sha256) userData.em = [data.email_sha256]
  else if (data.email) userData.em = [await sha256Hex(data.email)]
  if (data.phone_sha256) userData.ph = [data.phone_sha256]
  else if (data.phone) userData.ph = [await sha256Hex(data.phone.replace(/\D/g, ''))]
  // P1.4: fbp/fbc (cookies do Pixel, capturados no client em P1.1) e
  // client_user_agent melhoram o match quality da CAPI. Não bloqueante —
  // omitidos quando o checkout não trouxe esses campos (site ainda sem P1.1).
  if (data.fbp) userData.fbp = data.fbp
  if (data.fbc) userData.fbc = data.fbc
  if (data.client_user_agent) userData.client_user_agent = data.client_user_agent

  const body = {
    data: [{
      event_name: 'Purchase',
      // Hora real da conversão (paid_at); Meta CAPI aceita event_time até 7 dias atrás.
      event_time: Math.floor((data.paid_at ? new Date(data.paid_at).getTime() : Date.now()) / 1000),
      event_id: eventId,
      action_source: 'website',
      ...(data.event_source_url ? { event_source_url: data.event_source_url } : {}),
      user_data: userData,
      custom_data: {
        currency: 'BRL',
        value: data.valor,
        order_id: data.order_nsu,
      },
    }],
  }

  try {
    const resp = await fetch(
      `https://graph.facebook.com/v20.0/${pixelId}/events?access_token=${accessToken}`,
      { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
    )
    if (!resp.ok) return { status: 'failed', error: `Meta CAPI HTTP ${resp.status}: ${await resp.text()}` }
    return { status: 'success' }
  } catch (err) {
    return { status: 'failed', error: err instanceof Error ? err.message : String(err) }
  }
}

export async function dispatchAll(data: CheckoutConversionData): Promise<DispatchResult & { eventId: string }> {
  const eventId = computeEventId(data)
  const [ga4, ads, meta] = await Promise.all([
    dispatchGA4(data, eventId),
    dispatchGoogleAds(data),
    dispatchMeta(data, eventId),
  ])
  return { ga4, ads, meta, eventId }
}

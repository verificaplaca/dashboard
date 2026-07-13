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

async function sha256Hex(input: string): Promise<string> {
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

  // GA4 MP exige client_id. Não temos o _ga real (não é capturado no checkout hoje),
  // então geramos um synthetic determinístico a partir do checkout_id — o hit é registrado
  // no GA4 mas NÃO se junta à sessão original do usuário (limitação conhecida, não é bug).
  const clientId = `srv.${data.checkout_id.replace(/-/g, '').slice(0, 16)}`

  const body = {
    client_id: clientId,
    events: [{
      name: 'purchase',
      params: {
        transaction_id: data.order_nsu,
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

async function dispatchGoogleAds(data: CheckoutConversionData): Promise<DispatchResult['ads']> {
  const customerId       = Deno.env.get('GADS_CUSTOMER_ID')
  const devToken          = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const conversionActionId = Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID') // ID numérico — ainda não confirmado, ver doc de setup
  if (!customerId || !devToken || !conversionActionId) {
    return { status: 'skipped', error: 'GADS_PURCHASE_CONVERSION_ACTION_ID não configurado (ver instrucoes-setup-passos-5-7.md)' }
  }
  if (!data.email && !data.phone) {
    return { status: 'skipped', error: 'checkout sem email/phone (auth.users) — Enhanced Conversions for Leads exige ao menos um identificador' }
  }

  try {
    const accessToken = await getGadsAccessToken()
    const userIdentifiers: Record<string, unknown>[] = []
    if (data.email) userIdentifiers.push({ hashedEmail: await sha256Hex(data.email) })
    if (data.phone) userIdentifiers.push({ hashedPhoneNumber: await sha256Hex(data.phone) })

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
            orderId: data.order_nsu,
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
  if (data.email) userData.em = [await sha256Hex(data.email)]
  if (data.phone) userData.ph = [await sha256Hex(data.phone.replace(/\D/g, ''))]

  const body = {
    data: [{
      event_name: 'Purchase',
      event_time: Math.floor(Date.now() / 1000),
      event_id: eventId,
      action_source: 'website',
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

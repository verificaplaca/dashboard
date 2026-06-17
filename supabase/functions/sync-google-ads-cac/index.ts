/**
 * sync-google-ads-cac
 *
 * Busca search terms e keywords do Google Ads (com métricas de PURCHASE) e faz
 * upsert idempotente em google_ads_search_terms_daily / google_ads_keywords_daily.
 * Alimenta o módulo "Google Ads CAC" (google-ads-cac.html).
 *
 * Segue o mesmo padrão de supabase/functions/sync-google-ads/index.ts (campanhas),
 * adaptado para search_term_view + ad_group_criterion (keywords).
 *
 * Aceita body JSON opcional:
 *   { "start": "2026-01-01", "end": "2026-04-29" }  — backfill
 *   Sem body: últimos 3 dias (incremental)
 *
 * Secrets necessários (mesmos da sync-google-ads):
 *   GADS_CLIENT_ID       — OAuth client ID
 *   GADS_CLIENT_SECRET   — OAuth client secret
 *   GADS_REFRESH_TOKEN   — refresh token permanente
 *   GADS_DEVELOPER_TOKEN — developer token da MCC
 *   GADS_CUSTOMER_ID     — customer ID sem hífens (ex: 9258555135)
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const GADS_API_VERSION = 'v24'
const INTEGRATION       = 'google-ads-cac'

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function daysAgo(n: number): string {
  const d = new Date()
  d.setUTCDate(d.getUTCDate() - n)
  return d.toISOString().slice(0, 10)
}

async function getAccessToken(clientId: string, clientSecret: string, refreshToken: string): Promise<string> {
  const resp = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id:     clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type:    'refresh_token',
    }),
  })
  if (!resp.ok) throw new Error('OAuth token error: ' + await resp.text())
  const data = await resp.json() as { access_token: string }
  return data.access_token
}

async function gadsSearch(
  accessToken: string, devToken: string, customerId: string, query: string
): Promise<unknown[]> {
  const resp = await fetch(
    `https://googleads.googleapis.com/${GADS_API_VERSION}/customers/${customerId}/googleAds:search`,
    {
      method: 'POST',
      headers: {
        Authorization:      `Bearer ${accessToken}`,
        'developer-token':  devToken,
        'Content-Type':     'application/json',
      },
      body: JSON.stringify({ query }),
    }
  )
  if (!resp.ok) {
    const body = await resp.text()
    throw new Error(`Google Ads API HTTP ${resp.status}: ${body}`)
  }
  const data = await resp.json() as { results?: unknown[] }
  return data.results ?? []
}

type SearchTermRow = {
  segments:  { date: string }
  campaign:  { id: string; name: string }
  adGroup:   { id: string; name: string }
  searchTermView: { searchTerm: string }
  metrics:   { impressions: string; clicks: string; costMicros: string }
}

type KeywordRow = {
  segments:  { date: string }
  campaign:  { id: string; name: string }
  adGroup:   { id: string; name: string }
  adGroupCriterion: {
    keyword: { text: string; matchType: string }
    status: string
  }
  metrics:   { impressions: string; clicks: string; costMicros: string }
}

type PurchaseRow = {
  segments: { date: string }
  campaign: { id: string }
  adGroup:  { id: string }
  searchTermView?: { searchTerm: string }
  adGroupCriterion?: { keyword: { text: string } }
  metrics:  { conversions: string | number }
}

// Agrega conversões PURCHASE em um Map por chave date|campaignId|adGroupId|texto
function purchasesByKey(rows: PurchaseRow[], textKey: 'searchTermView' | 'adGroupCriterion'): Map<string, number> {
  const map = new Map<string, number>()
  for (const r of rows) {
    const text = textKey === 'searchTermView'
      ? r.searchTermView?.searchTerm
      : r.adGroupCriterion?.keyword?.text
    if (!text) continue
    const key = [r.segments.date, r.campaign.id, r.adGroup.id, text].join('|')
    const conv = parseFloat(String(r.metrics.conversions ?? '0')) || 0
    map.set(key, (map.get(key) ?? 0) + conv)
  }
  return map
}

Deno.serve(async (req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const clientId     = Deno.env.get('GADS_CLIENT_ID')
  const clientSecret  = Deno.env.get('GADS_CLIENT_SECRET')
  const refreshToken  = Deno.env.get('GADS_REFRESH_TOKEN')
  const devToken      = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId    = Deno.env.get('GADS_CUSTOMER_ID')

  if (!clientId || !clientSecret || !refreshToken || !devToken || !customerId) {
    return json({ ok: false, error: 'Secrets GADS_* não configurados' }, 500)
  }

  let start = daysAgo(3)
  let end   = todayISO()
  try {
    const body = await req.json().catch(() => null)
    if (body?.start) start = body.start
    if (body?.end)   end   = body.end
  } catch { /* ok */ }

  const { data: run } = await supabase
    .from('sync_runs')
    .insert({
      integration:  INTEGRATION,
      status:       'running',
      window_start: new Date(start + 'T00:00:00Z').toISOString(),
      window_end:   new Date(end   + 'T00:00:00Z').toISOString(),
    })
    .select('id')
    .single()

  const runId = run?.id

  try {
    const accessToken = await getAccessToken(clientId, clientSecret, refreshToken)
    const dateRange = `segments.date BETWEEN '${start}' AND '${end}'`

    // ── Search Terms: clicks/impressions/cost (todas conversões agregadas, não usado) ──
    const searchTermsRaw = await gadsSearch(accessToken, devToken, customerId, `
      SELECT
        segments.date,
        campaign.id, campaign.name,
        ad_group.id, ad_group.name,
        search_term_view.search_term,
        metrics.impressions, metrics.clicks, metrics.cost_micros
      FROM search_term_view
      WHERE ${dateRange}
        AND campaign.status != 'REMOVED'
      ORDER BY segments.date DESC
    `) as SearchTermRow[]

    // ── Search Terms: só conversões PURCHASE (mesmo critério do gads-intraday-script.js) ──
    const searchTermsPurchaseRaw = await gadsSearch(accessToken, devToken, customerId, `
      SELECT
        segments.date, segments.conversion_action_category,
        campaign.id, ad_group.id, search_term_view.search_term,
        metrics.conversions
      FROM search_term_view
      WHERE ${dateRange}
        AND campaign.status != 'REMOVED'
        AND segments.conversion_action_category = 'PURCHASE'
    `) as PurchaseRow[]
    const searchTermsPurchases = purchasesByKey(searchTermsPurchaseRaw, 'searchTermView')

    // ── Keywords: clicks/impressions/cost + status/match_type ──────────────────
    const keywordsRaw = await gadsSearch(accessToken, devToken, customerId, `
      SELECT
        segments.date,
        campaign.id, campaign.name,
        ad_group.id, ad_group.name,
        ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type,
        ad_group_criterion.status,
        metrics.impressions, metrics.clicks, metrics.cost_micros
      FROM keyword_view
      WHERE ${dateRange}
        AND campaign.status != 'REMOVED'
        AND ad_group_criterion.status != 'REMOVED'
      ORDER BY segments.date DESC
    `) as KeywordRow[]

    // ── Keywords: só conversões PURCHASE ────────────────────────────────────────
    const keywordsPurchaseRaw = await gadsSearch(accessToken, devToken, customerId, `
      SELECT
        segments.date, segments.conversion_action_category,
        campaign.id, ad_group.id, ad_group_criterion.keyword.text,
        metrics.conversions
      FROM keyword_view
      WHERE ${dateRange}
        AND campaign.status != 'REMOVED'
        AND segments.conversion_action_category = 'PURCHASE'
    `) as PurchaseRow[]
    const keywordsPurchases = purchasesByKey(keywordsPurchaseRaw, 'adGroupCriterion')

    // ── Keywords ativas (para is_existing_keyword nos search terms) ────────────
    const activeKeywordTexts = new Set(
      keywordsRaw
        .filter(r => r.adGroupCriterion.status === 'ENABLED')
        .map(r => r.adGroupCriterion.keyword.text.toLowerCase())
    )

    // ── Transforma search terms ──────────────────────────────────────────────
    const searchTermRows = searchTermsRaw.map(r => {
      const key = [r.segments.date, r.campaign.id, r.adGroup.id, r.searchTermView.searchTerm].join('|')
      return {
        date:                r.segments.date,
        search_term:         r.searchTermView.searchTerm,
        campaign_id:         r.campaign.id,
        campaign_name:       r.campaign.name,
        ad_group_id:         r.adGroup.id,
        ad_group_name:       r.adGroup.name,
        clicks:              parseInt(r.metrics.clicks ?? '0') || 0,
        impressions:         parseInt(r.metrics.impressions ?? '0') || 0,
        cost_micros:         parseInt(r.metrics.costMicros ?? '0') || 0,
        purchases:           searchTermsPurchases.get(key) ?? 0,
        is_existing_keyword: activeKeywordTexts.has(r.searchTermView.searchTerm.toLowerCase()),
        raw_json:            r,
        ingested_at:         new Date().toISOString(),
      }
    })

    // ── Transforma keywords ──────────────────────────────────────────────────
    const matchTypeMap: Record<string, string> = { BROAD: 'broad', PHRASE: 'phrase', EXACT: 'exact' }
    const keywordRows = keywordsRaw.map(r => {
      const key = [r.segments.date, r.campaign.id, r.adGroup.id, r.adGroupCriterion.keyword.text].join('|')
      return {
        date:               r.segments.date,
        keyword:            r.adGroupCriterion.keyword.text,
        match_type:         matchTypeMap[r.adGroupCriterion.keyword.matchType] ?? 'broad',
        campaign_id:        r.campaign.id,
        campaign_name:      r.campaign.name,
        ad_group_id:        r.adGroup.id,
        ad_group_name:      r.adGroup.name,
        status_google_ads:  r.adGroupCriterion.status,
        clicks:             parseInt(r.metrics.clicks ?? '0') || 0,
        impressions:        parseInt(r.metrics.impressions ?? '0') || 0,
        cost_micros:        parseInt(r.metrics.costMicros ?? '0') || 0,
        purchases:          keywordsPurchases.get(key) ?? 0,
        raw_json:           r,
        ingested_at:        new Date().toISOString(),
      }
    })

    // ── Upsert idempotente (requer UNIQUE INDEX — ver google_ads_cac_sync.sql) ──
    if (searchTermRows.length > 0) {
      const { error } = await supabase
        .from('google_ads_search_terms_daily')
        .upsert(searchTermRows, { onConflict: 'date,search_term,campaign_id,ad_group_id' })
      if (error) throw new Error('search_terms upsert: ' + error.message)
    }

    if (keywordRows.length > 0) {
      const { error } = await supabase
        .from('google_ads_keywords_daily')
        .upsert(keywordRows, { onConflict: 'date,keyword,campaign_id,ad_group_id' })
      if (error) throw new Error('keywords upsert: ' + error.message)
    }

    const totalRecords = searchTermRows.length + keywordRows.length

    await supabase.from('sync_runs').update({
      status:            'completed',
      records_processed: totalRecords,
      finished_at:        new Date().toISOString(),
    }).eq('id', runId)

    return json({
      ok: true,
      search_terms: searchTermRows.length,
      keywords: keywordRows.length,
      window: `${start} → ${end}`,
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

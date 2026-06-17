/**
 * sync-google-ads
 *
 * Busca performance diária de campanhas Google Ads e faz upsert
 * idempotente em google_ads_campaign_daily.
 *
 * Aceita body JSON opcional:
 *   { "start": "2026-01-01", "end": "2026-04-29" }  — backfill
 *   Sem body: últimos 3 dias (incremental)
 *
 * Secrets necessários:
 *   GADS_CLIENT_ID       — OAuth client ID
 *   GADS_CLIENT_SECRET   — OAuth client secret
 *   GADS_REFRESH_TOKEN   — refresh token permanente
 *   GADS_DEVELOPER_TOKEN — developer token da MCC
 *   GADS_CUSTOMER_ID     — customer ID sem hífens (ex: 9258555135)
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const GADS_API_VERSION = 'v24'
const INTEGRATION      = 'google-ads'

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

Deno.serve(async (req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const clientId      = Deno.env.get('GADS_CLIENT_ID')
  const clientSecret  = Deno.env.get('GADS_CLIENT_SECRET')
  const refreshToken  = Deno.env.get('GADS_REFRESH_TOKEN')
  const devToken      = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId    = Deno.env.get('GADS_CUSTOMER_ID')

  if (!clientId || !clientSecret || !refreshToken || !devToken || !customerId) {
    return json({ ok: false, error: 'Secrets GADS_* não configurados' }, 500)
  }

  // Lê datas do body ou usa padrão (últimos 3 dias)
  let start = daysAgo(3)
  let end   = todayISO()
  try {
    const body = await req.json().catch(() => null)
    if (body?.start) start = body.start
    if (body?.end)   end   = body.end
  } catch { /* ok */ }

  // Registra início do run
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

    // ── Consulta GAQL ─────────────────────────────────────────
    const query = `
      SELECT
        campaign.id,
        campaign.name,
        metrics.impressions,
        metrics.clicks,
        metrics.cost_micros,
        metrics.conversions,
        segments.date
      FROM campaign
      WHERE segments.date BETWEEN '${start}' AND '${end}'
      ORDER BY segments.date DESC, metrics.cost_micros DESC
    `

    const gadsResp = await fetch(
      `https://googleads.googleapis.com/${GADS_API_VERSION}/customers/${customerId}/googleAds:search`,
      {
        method: 'POST',
        headers: {
          Authorization:           `Bearer ${accessToken}`,
          'developer-token':        devToken,
          'Content-Type':           'application/json',
        },
        body: JSON.stringify({ query }),
      }
    )

    if (!gadsResp.ok) {
      const body = await gadsResp.text()
      throw new Error(`Google Ads API HTTP ${gadsResp.status}: ${body}`)
    }

    const gadsData = await gadsResp.json() as { results?: unknown[] }
    const results  = gadsData.results ?? []

    if (results.length === 0) {
      await supabase.from('sync_runs').update({
        status: 'completed', records_processed: 0, finished_at: new Date().toISOString(),
      }).eq('id', runId)
      return json({ ok: true, records: 0, window: `${start} → ${end}` })
    }

    // ── Transforma e faz upsert ───────────────────────────────
    const rows = (results as Array<{
      campaign:  { id: string; name: string }
      metrics:   { impressions: string; clicks: string; costMicros: string; conversions: string | number }
      segments:  { date: string }
    }>).map(r => ({
      date:          r.segments.date,
      campaign_id:   r.campaign.id,
      campaign_name: r.campaign.name,
      impressions:   parseInt(r.metrics.impressions ?? '0') || 0,
      clicks:        parseInt(r.metrics.clicks       ?? '0') || 0,
      cost_micros:   parseInt(r.metrics.costMicros   ?? '0') || 0,
      conversions:   parseFloat(String(r.metrics.conversions ?? '0')) || 0,
      raw_json:      r,
      ingested_at:   new Date().toISOString(),
    }))

    const { error: upsertErr } = await supabase
      .from('google_ads_campaign_daily')
      .upsert(rows, { onConflict: 'date,campaign_id' })

    if (upsertErr) throw new Error('google_ads upsert: ' + upsertErr.message)

    await supabase.from('sync_runs').update({
      status:            'completed',
      records_processed: rows.length,
      finished_at:       new Date().toISOString(),
    }).eq('id', runId)

    return json({ ok: true, records: rows.length, window: `${start} → ${end}` })

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

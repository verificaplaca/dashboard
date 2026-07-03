/**
 * sync-bureau-daily
 *
 * Busca custo diário de bureau no Supabase externo (Assertiva / checkouts)
 * e faz upsert idempotente em bureau_daily.
 *
 * Aceita body JSON opcional:
 *   { "p_start": "2026-01-01", "p_end": "2026-04-30" }  — para backfill
 *   Sem body: últimos 3 dias (incremental diário)
 *
 * Secrets necessários:
 *   BUREAU_SUPABASE_URL — https://ozquoloetuzynnyzkado.supabase.co
 *   BUREAU_SUPABASE_KEY — anon key do Supabase externo
 *
 * Deploy:  supabase functions deploy sync-bureau-daily --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:    cron-job.org → POST .../functions/v1/sync-bureau-daily, 3x/dia (06h, 12h, 18h BRT)
 *          Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 *
 * Depende da função RPC get_bureau_costs_daily(p_start, p_end), definida em
 * supabase/bureau_daily.sql — mas rodando no projeto EXTERNO (BUREAU_SUPABASE_URL),
 * não neste projeto.
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const INTEGRATION = 'bureau-daily'

function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

function daysAgo(n: number): string {
  const d = new Date()
  d.setUTCDate(d.getUTCDate() - n)
  return d.toISOString().slice(0, 10)
}

Deno.serve(async (req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const bureauUrl = Deno.env.get('BUREAU_SUPABASE_URL')
  const bureauKey = Deno.env.get('BUREAU_SUPABASE_KEY')

  if (!bureauUrl || !bureauKey) {
    return json({ ok: false, error: 'BUREAU_SUPABASE_URL ou BUREAU_SUPABASE_KEY não configurados' }, 500)
  }

  // Lê datas do body (backfill) ou usa padrão incremental (últimos 3 dias)
  let p_start = daysAgo(3)
  let p_end   = todayISO()
  try {
    const body = await req.json().catch(() => null)
    if (body?.p_start) p_start = body.p_start
    if (body?.p_end)   p_end   = body.p_end
  } catch { /* ok */ }

  // Registra início do run
  const { data: run } = await supabase
    .from('sync_runs')
    .insert({
      integration:  INTEGRATION,
      status:       'running',
      window_start: new Date(p_start + 'T00:00:00Z').toISOString(),
      window_end:   new Date(p_end   + 'T00:00:00Z').toISOString(),
    })
    .select('id')
    .single()

  const runId = run?.id

  try {
    // ── Chama RPC no Supabase externo ─────────────────────────
    const resp = await fetch(`${bureauUrl}/rest/v1/rpc/get_bureau_costs_daily`, {
      method: 'POST',
      headers: {
        apikey:        bureauKey,
        Authorization: `Bearer ${bureauKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ p_start, p_end }),
    })

    if (!resp.ok) {
      const body = await resp.text()
      throw new Error(`Bureau RPC HTTP ${resp.status}: ${body}`)
    }

    const rows = await resp.json() as Array<{
      dia:          string
      vendas_pagas: number
      vendido_real: number
      custo_bureau: number
      lucro_bruto:  number
      margem_pct:   number
    }>

    if (!Array.isArray(rows)) {
      throw new Error('Resposta inesperada do bureau: ' + JSON.stringify(rows))
    }

    if (rows.length === 0) {
      await supabase.from('sync_runs').update({
        status: 'completed', records_processed: 0, finished_at: new Date().toISOString(),
      }).eq('id', runId)
      return json({ ok: true, records: 0 })
    }

    // ── Upsert em bureau_daily ────────────────────────────────
    const upsertRows = rows.map(r => ({
      date:         r.dia,
      vendas_pagas: r.vendas_pagas,
      vendido_real: r.vendido_real,
      custo_bureau: r.custo_bureau,
      lucro_bruto:  r.lucro_bruto,
      margem_pct:   r.margem_pct,
      ingested_at:  new Date().toISOString(),
    }))

    const { error: upsertErr } = await supabase
      .from('bureau_daily')
      .upsert(upsertRows, { onConflict: 'date' })

    if (upsertErr) throw new Error('bureau_daily upsert: ' + upsertErr.message)

    await supabase.from('sync_runs').update({
      status:            'completed',
      records_processed: rows.length,
      finished_at:       new Date().toISOString(),
    }).eq('id', runId)

    return json({ ok: true, records: rows.length })

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

/**
 * sync-bureau-by-type
 *
 * Edge Function NOVA, em paralelo a sync-bureau-daily — não altera a existente.
 *
 * Busca custo diário de bureau QUEBRADO POR TIPO (Assertiva vs CheckTudo) no
 * Supabase do Site/Sistema e faz upsert idempotente em bureau_by_type_daily
 * (Dashboard). Objetivo: acompanhar quanto cada bureau custa enquanto os dois
 * ficam ligados em paralelo.
 *
 * Aceita body JSON opcional:
 *   { "p_start": "2026-01-01", "p_end": "2026-04-30" }  — para backfill
 *   Sem body: últimos 35 dias — janela larga de propósito: a bureau_chamadas_cobradas
 *   lança/ajusta cobranças RETROATIVAMENTE (com dias de atraso); janela curta
 *   congelava dias antigos com custo subestimado (bug corrigido em 2026-07-16)
 *
 * Secrets necessários (iguais aos de sync-bureau-daily):
 *   BUREAU_SUPABASE_URL — https://ozquoloetuzynnyzkado.supabase.co (Site/Sistema)
 *   BUREAU_SUPABASE_KEY — anon key do Supabase do Site/Sistema
 *
 * Deploy:  supabase functions deploy sync-bureau-by-type --project-ref ftmgmfdqdqxboiktxcoj (Dashboard)
 * Cron:    cron-job.org → POST .../functions/v1/sync-bureau-by-type, mesma cadência
 *          de sync-bureau-daily (a cada 30min desde 16/07/2026)
 *          Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 *
 * Depende da função RPC get_bureau_costs_by_bureau(p_start, p_end), definida em
 * supabase/bureau_by_type.sql — rodando no Supabase do Site/Sistema, não neste
 * projeto (Dashboard).
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const INTEGRATION = 'bureau-by-type-daily'

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

  let p_start = daysAgo(35)
  let p_end   = todayISO()
  try {
    const body = await req.json().catch(() => null)
    if (body?.p_start) p_start = body.p_start
    if (body?.p_end)   p_end   = body.p_end
  } catch { /* ok */ }

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
    const resp = await fetch(`${bureauUrl}/rest/v1/rpc/get_bureau_costs_by_bureau`, {
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
      bureau:       string
      vendas_pagas: number
      vendido_real: number
      custo_bureau: number
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

    const upsertRows = rows.map(r => ({
      date:         r.dia,
      bureau:       r.bureau,
      vendas_pagas: r.vendas_pagas,
      vendido_real: r.vendido_real,
      custo_bureau: r.custo_bureau,
      ingested_at:  new Date().toISOString(),
    }))

    const { error: upsertErr } = await supabase
      .from('bureau_by_type_daily')
      .upsert(upsertRows, { onConflict: 'date,bureau' })

    if (upsertErr) throw new Error('bureau_by_type_daily upsert: ' + upsertErr.message)

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

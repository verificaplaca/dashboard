/**
 * check-tracking-health — Alerta ativo de saúde do tracking (P2.2) + watchdog de sync
 *
 * Últimas 24h:
 *   (a) % de linhas de conversion_dispatches com QUALQUER canal 'failed' > 10%
 *   (b) conversion_dispatches = 0 linhas enquanto revenue_daily.paid_orders > 0
 *       no dia de hoje (indício de webhook morto)
 *
 * Watchdog de frescor (novos checks):
 *   (c) bureau_daily.ingested_at mais recente > 3h atrás (sync roda a cada 30min)
 *   (d) ads_balance_history.checked_at mais recente > 2h atrás (check roda 1x/h)
 *   (e) google_ads_campaign_daily sem linha para hoje (BRT), só checado depois
 *       das 10:00 BRT (antes disso o primeiro sync do dia pode não ter rodado)
 *   (f) sync_runs das integrations 'bureau-daily' e 'bureau-by-type-daily':
 *       último run com status='failed', OU últimos 6 runs todos com
 *       records_processed=0 (sintoma do incidente de RLS de 14/07 — a RPC
 *       retorna vazio sem erro; o sync sempre cobre 3 dias, então 0 records
 *       repetido = quebrado)
 *
 * Todos os problemas achados em uma execução são agrupados numa única
 * mensagem no Telegram. Dedup via `health_alerts`: cada tipo de problema tem
 * um alert_key estável; se já foi enviado nas últimas 6h, não repete.
 *
 * Variáveis de ambiente necessárias:
 *   TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID  (já existem, reaproveitados)
 *   SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY  (injetados automaticamente)
 *
 * Deploy: supabase functions deploy check-tracking-health --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:   cron-job.org → POST .../functions/v1/check-tracking-health, 1x/hora
 *         Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const FAILED_RATE_THRESHOLD = 0.10
const DEDUP_WINDOW_HOURS = 6
const BUREAU_SYNC_INTEGRATIONS = ['bureau-daily', 'bureau-by-type-daily']

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } })
}

async function sendTelegram(botToken: string, chatId: string, text: string): Promise<void> {
  const resp = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: 'HTML' }),
  })
  if (!resp.ok) console.error('Telegram error:', await resp.text())
}

// ── BRT (America/Sao_Paulo, UTC-3 fixo) ─────────────────────────────────────
function todayBRT(): string {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo' }).format(new Date())
}

function currentHourBRT(): number {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Sao_Paulo', hour: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date())
  return Number(parts.find(p => p.type === 'hour')?.value ?? '0')
}

function hoursAgo(iso: string | null | undefined): number | null {
  if (!iso) return null
  return (Date.now() - new Date(iso).getTime()) / 3600_000
}

Deno.serve(async () => {
  try {
    const supaUrl  = Deno.env.get('SUPABASE_URL')
    const supaKey  = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
    const botToken = Deno.env.get('TELEGRAM_BOT_TOKEN')
    const chatId   = Deno.env.get('TELEGRAM_CHAT_ID')

    const missing = [
      !supaUrl  && 'SUPABASE_URL',
      !supaKey  && 'SUPABASE_SERVICE_ROLE_KEY',
      !botToken && 'TELEGRAM_BOT_TOKEN',
      !chatId   && 'TELEGRAM_CHAT_ID',
    ].filter(Boolean)
    if (missing.length) return json({ error: `Missing env vars: ${missing.join(', ')}` }, 500)

    const supabase = createClient(supaUrl!, supaKey!)

    const since24h = new Date(Date.now() - 24 * 3600_000).toISOString()
    const todayStr = new Date().toISOString().slice(0, 10)

    // ── (a) % failed nas últimas 24h ──────────────────────────────────────
    const { data: rows, error: rowsErr } = await supabase
      .from('conversion_dispatches')
      .select('ga4_status, ads_status, meta_status')
      .gte('created_at', since24h)
    if (rowsErr) return json({ error: `conversion_dispatches select: ${rowsErr.message}` }, 500)

    const total = rows?.length ?? 0
    const failedCount = (rows ?? []).filter(
      r => r.ga4_status === 'failed' || r.ads_status === 'failed' || r.meta_status === 'failed'
    ).length
    const failedRate = total > 0 ? failedCount / total : 0
    const highFailedRate = total > 0 && failedRate > FAILED_RATE_THRESHOLD

    // ── (b) webhook morto: 0 dispatches hoje com paid_orders > 0 hoje ─────
    const { data: revRows, error: revErr } = await supabase
      .from('revenue_daily')
      .select('date, paid_orders')
      .eq('date', todayStr)
    if (revErr) return json({ error: `revenue_daily select: ${revErr.message}` }, 500)

    const paidOrdersToday = (revRows ?? []).reduce((s, r) => s + (Number(r.paid_orders) || 0), 0)

    const { count: dispatchesToday, error: todayErr } = await supabase
      .from('conversion_dispatches')
      .select('checkout_id', { count: 'exact', head: true })
      .gte('created_at', `${todayStr}T00:00:00Z`)
    if (todayErr) return json({ error: `conversion_dispatches count: ${todayErr.message}` }, 500)

    const webhookDead = paidOrdersToday > 0 && (dispatchesToday ?? 0) === 0

    // ── Problemas detectados nesta execução (alert_key + linha da mensagem) ─
    const problems: { key: string; line: string }[] = []

    if (highFailedRate) {
      problems.push({
        key: 'dispatch_failed_rate',
        line: `🔴 Taxa de falha alta: <b>${(failedRate * 100).toFixed(1)}%</b> (${failedCount}/${total}) nas últimas 24h`,
      })
    }
    if (webhookDead) {
      problems.push({
        key: 'dispatch_zero',
        line: `🔴 Zero linhas em conversion_dispatches hoje, mas <b>${paidOrdersToday} pedidos pagos</b> em revenue_daily — possível webhook morto`,
      })
    }

    // ── (c) bureau_daily — frescor (ingested_at > 3h) ──────────────────────
    const { data: bureauFreshRows, error: bureauFreshErr } = await supabase
      .from('bureau_daily')
      .select('ingested_at')
      .order('ingested_at', { ascending: false })
      .limit(1)
    if (bureauFreshErr) return json({ error: `bureau_daily select: ${bureauFreshErr.message}` }, 500)

    const bureauStaleHours = hoursAgo(bureauFreshRows?.[0]?.ingested_at)
    const bureauStale = bureauStaleHours === null || bureauStaleHours > 3
    if (bureauStale) {
      problems.push({
        key: 'bureau_daily_stale',
        line: bureauStaleHours !== null
          ? `🟠 bureau_daily sem atualização há <b>${bureauStaleHours.toFixed(1)}h</b> (esperado a cada 30min)`
          : `🟠 bureau_daily sem nenhum registro`,
      })
    }

    // ── (d) ads_balance_history — frescor (checked_at > 2h) ────────────────
    const { data: balanceFreshRows, error: balanceFreshErr } = await supabase
      .from('ads_balance_history')
      .select('checked_at')
      .order('checked_at', { ascending: false })
      .limit(1)
    if (balanceFreshErr) return json({ error: `ads_balance_history select: ${balanceFreshErr.message}` }, 500)

    const balanceStaleHours = hoursAgo(balanceFreshRows?.[0]?.checked_at)
    const balanceStale = balanceStaleHours === null || balanceStaleHours > 2
    if (balanceStale) {
      problems.push({
        key: 'ads_balance_stale',
        line: balanceStaleHours !== null
          ? `🟠 ads_balance_history sem atualização há <b>${balanceStaleHours.toFixed(1)}h</b> (esperado a cada 1h)`
          : `🟠 ads_balance_history sem nenhum registro`,
      })
    }

    // ── (e) google_ads_campaign_daily — sem linha de hoje (BRT), só após 10h ─
    const todayBRTStr = todayBRT()
    const hourBRT = currentHourBRT()
    if (hourBRT >= 10) {
      const { data: gadsTodayRows, error: gadsTodayErr } = await supabase
        .from('google_ads_campaign_daily')
        .select('date')
        .eq('date', todayBRTStr)
        .limit(1)
      if (gadsTodayErr) return json({ error: `google_ads_campaign_daily select: ${gadsTodayErr.message}` }, 500)

      if (!gadsTodayRows || gadsTodayRows.length === 0) {
        problems.push({
          key: 'gads_daily_missing',
          line: `🟠 google_ads_campaign_daily sem nenhuma linha para hoje (${todayBRTStr}) — já passou das 10h BRT`,
        })
      }
    }

    // ── (f) sync_runs — bureau-daily / bureau-by-type-daily quebrados ──────
    let syncBureauBroken = false
    const syncBureauDetails: string[] = []
    for (const integration of BUREAU_SYNC_INTEGRATIONS) {
      const { data: runs, error: runsErr } = await supabase
        .from('sync_runs')
        .select('status, records_processed')
        .eq('integration', integration)
        .in('status', ['completed', 'failed'])
        .order('started_at', { ascending: false })
        .limit(6)
      if (runsErr) return json({ error: `sync_runs select (${integration}): ${runsErr.message}` }, 500)

      if (!runs || runs.length === 0) continue

      const lastFailed = runs[0].status === 'failed'
      const allZero = runs.length >= 6 && runs.every(r => Number(r.records_processed ?? 0) === 0)

      if (lastFailed || allZero) {
        syncBureauBroken = true
        syncBureauDetails.push(
          lastFailed
            ? `${integration}: último run falhou`
            : `${integration}: últimos ${runs.length} runs com 0 registros processados`,
        )
      }
    }
    if (syncBureauBroken) {
      problems.push({
        key: 'sync_bureau_failed',
        line: `🔴 Sync de bureau quebrado — ${syncBureauDetails.join('; ')}`,
      })
    }

    // ── Dedup: pula alert_key já enviado nas últimas ${DEDUP_WINDOW_HOURS}h ─
    const sinceDedup = new Date(Date.now() - DEDUP_WINDOW_HOURS * 3600_000).toISOString()
    const { data: recentAlerts, error: recentAlertsErr } = await supabase
      .from('health_alerts')
      .select('alert_key')
      .gte('sent_at', sinceDedup)
    if (recentAlertsErr) return json({ error: `health_alerts select: ${recentAlertsErr.message}` }, 500)

    const recentKeys = new Set((recentAlerts ?? []).map(r => r.alert_key))
    const newProblems = problems.filter(p => !recentKeys.has(p.key))

    let alertSent = false
    if (newProblems.length > 0) {
      const lines = [`🚨 <b>Alerta de saúde do tracking</b>`, ``, ...newProblems.map(p => p.line)]
      lines.push('', 'Verificar dispatch-conversions / retry-conversion-dispatches / sync-bureau-daily / sync-bureau-by-type.')
      await sendTelegram(botToken!, chatId!, lines.join('\n'))
      alertSent = true

      await supabase.from('health_alerts').insert(
        newProblems.map(p => ({ alert_key: p.key }))
      )
    }

    return json({
      ok: true,
      total_24h: total,
      failed_count_24h: failedCount,
      failed_rate_24h: failedRate,
      high_failed_rate: highFailedRate,
      paid_orders_today: paidOrdersToday,
      dispatches_today: dispatchesToday ?? 0,
      webhook_dead: webhookDead,
      bureau_daily_stale: bureauStale,
      ads_balance_stale: balanceStale,
      sync_bureau_broken: syncBureauBroken,
      problems_detected: problems.map(p => p.key),
      problems_alerted: newProblems.map(p => p.key),
      alert_sent: alertSent,
      checked_at: new Date().toISOString(),
    })
  } catch (err) {
    return json({ error: String(err) }, 500)
  }
})

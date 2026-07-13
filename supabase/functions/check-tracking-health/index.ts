/**
 * check-tracking-health — Alerta ativo de saúde do tracking (P2.2)
 *
 * Últimas 24h:
 *   (a) % de linhas de conversion_dispatches com QUALQUER canal 'failed' > 10%
 *   (b) conversion_dispatches = 0 linhas enquanto revenue_daily.paid_orders > 0
 *       no dia de hoje (indício de webhook morto)
 * Qualquer uma das duas condições dispara alerta no Telegram (mesmo mecanismo
 * de notificação da check-ads-balance — reaproveita TELEGRAM_BOT_TOKEN/
 * TELEGRAM_CHAT_ID já configurados nesta conta).
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

    const alertNeeded = highFailedRate || webhookDead
    let alertSent = false

    if (alertNeeded) {
      const lines = [`🚨 <b>Alerta de saúde do tracking</b>`, ``]
      if (highFailedRate) {
        lines.push(
          `🔴 Taxa de falha alta: <b>${(failedRate * 100).toFixed(1)}%</b> (${failedCount}/${total}) nas últimas 24h`,
        )
      }
      if (webhookDead) {
        lines.push(
          `🔴 Zero linhas em conversion_dispatches hoje, mas <b>${paidOrdersToday} pedidos pagos</b> em revenue_daily — possível webhook morto`,
        )
      }
      lines.push('', 'Verificar dispatch-conversions / retry-conversion-dispatches.')
      await sendTelegram(botToken!, chatId!, lines.join('\n'))
      alertSent = true
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
      alert_sent: alertSent,
      checked_at: new Date().toISOString(),
    })
  } catch (err) {
    return json({ error: String(err) }, 500)
  }
})

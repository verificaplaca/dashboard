/**
 * resumo-diario — Resumo diário do dia anterior → Telegram
 *
 * Roda 1x/dia (cron 08:00 BRT) e manda um resumo compacto de ONTEM (data
 * calculada em America/Sao_Paulo, não UTC) com receita, upsell, reembolsos,
 * custos de ads/bureau, lucro bruto/líquido, CAC real, saldo de Ads e saúde
 * do tracking. Fontes de dados (tabelas do próprio projeto, service role):
 *   revenue_daily, upsell_daily, bureau_daily, google_ads_campaign_daily,
 *   refunds_daily, ads_balance_history, conversion_dispatches
 *
 * Sem tabela de dedup — o cron roda só 1x/dia, não precisa. Se alguma fonte
 * não tiver dado de ontem, a linha correspondente vira "—" e a fonte entra
 * na lista de "sem dados" no fim da mensagem (não falha a mensagem inteira).
 *
 * Cálculos (convenções do dashboard, CLAUDE.md):
 *   custo_total  = custo_ads + custo_bureau
 *   lucro_bruto  = receita − custo_total
 *   lucro_líquido = receita * 0.92 − custo_total
 *   CAC real     = custo_ads / paid_orders
 *
 * Variáveis de ambiente necessárias:
 *   TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID     (já existem, reaproveitados)
 *   SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY  (injetados automaticamente)
 *
 * Deploy: supabase functions deploy resumo-diario --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:   cron-job.org → POST .../functions/v1/resumo-diario, 1x/dia às 08:00 BRT (11:00 UTC)
 *         Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } })
}

function fmtBRL(value: number): string {
  return value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
}

async function sendTelegram(botToken: string, chatId: string, text: string): Promise<void> {
  const resp = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: 'HTML' }),
  })
  if (!resp.ok) console.error('Telegram error:', await resp.text())
}

// ── Datas em America/Sao_Paulo (BRT = UTC-3, sem horário de verão) ─────────
function todayBRT(): string {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo' }).format(new Date())
}

function yesterdayOf(dateStr: string): string {
  const d = new Date(dateStr + 'T00:00:00Z')
  d.setUTCDate(d.getUTCDate() - 1)
  return d.toISOString().slice(0, 10)
}

function daysBefore(dateStr: string, n: number): string {
  const d = new Date(dateStr + 'T00:00:00Z')
  d.setUTCDate(d.getUTCDate() - n)
  return d.toISOString().slice(0, 10)
}

Deno.serve(async () => {
  try {
    const supaUrl  = Deno.env.get('SUPABASE_URL')
    const supaKey  = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
    const botToken = Deno.env.get('TELEGRAM_BOT_TOKEN')
    const chatId   = Deno.env.get('TELEGRAM_CHAT_ID')

    const missingEnv = [
      !supaUrl  && 'SUPABASE_URL',
      !supaKey  && 'SUPABASE_SERVICE_ROLE_KEY',
      !botToken && 'TELEGRAM_BOT_TOKEN',
      !chatId   && 'TELEGRAM_CHAT_ID',
    ].filter(Boolean)
    if (missingEnv.length) return json({ error: `Missing env vars: ${missingEnv.join(', ')}` }, 500)

    const supabase = createClient(supaUrl!, supaKey!)

    const todayStr     = todayBRT()
    const yesterdayStr = yesterdayOf(todayStr)
    // Janela [ontem 00:00 BRT, hoje 00:00 BRT) em UTC — BRT = UTC-3, fixo (sem DST).
    const sinceUTC = `${yesterdayStr}T03:00:00.000Z`
    const untilUTC = `${todayStr}T03:00:00.000Z`

    const missing: string[] = []

    // ── revenue_daily ──────────────────────────────────────────────────────
    let revenue: number | null = null
    let paidOrders: number | null = null
    {
      const { data, error } = await supabase
        .from('revenue_daily')
        .select('date, revenue, paid_orders')
        .eq('date', yesterdayStr)
      if (error || !data || data.length === 0) {
        missing.push('revenue_daily')
      } else {
        revenue    = data.reduce((s, r) => s + (Number(r.revenue) || 0), 0)
        paidOrders = data.reduce((s, r) => s + (Number(r.paid_orders) || 0), 0)
      }
    }

    // ── upsell_daily ───────────────────────────────────────────────────────
    let upsellOrders: number | null = null
    let upsellRate: number | null = null
    {
      const { data, error } = await supabase
        .from('upsell_daily')
        .select('date, upsell_orders, upsell_rate')
        .eq('date', yesterdayStr)
      if (error || !data || data.length === 0) {
        missing.push('upsell_daily')
      } else {
        upsellOrders = data.reduce((s, r) => s + (Number(r.upsell_orders) || 0), 0)
        upsellRate   = Number(data[0].upsell_rate) || 0
      }
    }

    // ── refunds_daily ──────────────────────────────────────────────────────
    // Dia sem estorno não gera linha na view → ausência = zero (não é "sem dados").
    let refundCount = 0
    let refundValue = 0
    {
      const { data, error } = await supabase
        .from('refunds_daily')
        .select('date, refund_count, refund_value')
        .eq('date', yesterdayStr)
      if (error) {
        missing.push('refunds_daily')
      } else if (data && data.length > 0) {
        refundCount = data.reduce((s, r) => s + (Number(r.refund_count) || 0), 0)
        refundValue = data.reduce((s, r) => s + (Number(r.refund_value) || 0), 0)
      }
    }

    // ── bureau_daily ───────────────────────────────────────────────────────
    let custoBureau: number | null = null
    {
      const { data, error } = await supabase
        .from('bureau_daily')
        .select('date, custo_bureau')
        .eq('date', yesterdayStr)
      if (error || !data || data.length === 0) {
        missing.push('bureau_daily')
      } else {
        custoBureau = data.reduce((s, r) => s + (Number(r.custo_bureau) || 0), 0)
      }
    }

    // ── google_ads_campaign_daily (custo de ontem + burn rate 7d) ──────────
    let custoAds: number | null = null
    let avg7spend = 0
    {
      const since7 = daysBefore(yesterdayStr, 7)
      const { data, error } = await supabase
        .from('google_ads_campaign_daily')
        .select('date, cost_micros')
        .gte('date', since7)
        .lte('date', yesterdayStr)
      if (error || !data || data.length === 0) {
        missing.push('google_ads_campaign_daily')
      } else {
        const byDate = new Map<string, number>()
        for (const r of data) {
          const d = String(r.date).slice(0, 10)
          byDate.set(d, (byDate.get(d) ?? 0) + Number(r.cost_micros ?? 0))
        }
        custoAds = (byDate.get(yesterdayStr) ?? 0) / 1_000_000
        const dailyCosts = [...byDate.values()].map(m => m / 1_000_000).filter(v => v > 0)
        avg7spend = dailyCosts.length ? dailyCosts.reduce((s, v) => s + v, 0) / dailyCosts.length : 0
      }
    }

    // ── ads_balance_history (último snapshot, saldo + dias restantes) ──────
    let balanceBRL: number | null = null
    let daysLeft: number | null = null
    {
      const { data, error } = await supabase
        .from('ads_balance_history')
        .select('balance_brl, checked_at')
        .order('checked_at', { ascending: false })
        .limit(1)
      if (error || !data || data.length === 0) {
        missing.push('ads_balance_history')
      } else {
        balanceBRL = Number(data[0].balance_brl) || 0
        daysLeft   = avg7spend > 0 ? balanceBRL / avg7spend : null
      }
    }

    // ── conversion_dispatches (linhas de ontem) ────────────────────────────
    let dispatchTotal: number | null = null
    let dispatchFailed: number | null = null
    {
      const { data, error } = await supabase
        .from('conversion_dispatches')
        .select('ga4_status, ads_status, meta_status')
        .gte('created_at', sinceUTC)
        .lt('created_at', untilUTC)
      if (error) {
        missing.push('conversion_dispatches')
      } else {
        dispatchTotal  = data?.length ?? 0
        dispatchFailed = (data ?? []).filter(
          r => r.ga4_status === 'failed' || r.ads_status === 'failed' || r.meta_status === 'failed'
        ).length
      }
    }

    // ── Cálculos derivados ──────────────────────────────────────────────────
    const custoTotal = (custoAds ?? 0) + (custoBureau ?? 0)
    const custoTotalKnown = custoAds !== null || custoBureau !== null
    const lucroBruto  = revenue !== null && custoTotalKnown ? revenue - custoTotal : null
    const lucroLiquido = revenue !== null && custoTotalKnown ? revenue * 0.92 - custoTotal : null
    const margem = lucroBruto !== null && revenue ? lucroBruto / revenue : null
    const cacReal = custoAds !== null && paidOrders ? custoAds / paidOrders : null
    const ticketMedio = revenue !== null && paidOrders ? revenue / paidOrders : null

    // ── Acumulado do mês (mês de "ontem" em BRT, do dia 01 até ontem) ───────
    // COGS do mês = Ads + Bureau (custo total, convenção do dashboard).
    const monthStart = yesterdayStr.slice(0, 7) + '-01'
    let moRevenue: number | null = null
    let moAds: number | null = null
    let moBureau: number | null = null
    {
      const { data, error } = await supabase
        .from('revenue_daily')
        .select('date, revenue')
        .gte('date', monthStart)
        .lte('date', yesterdayStr)
      if (error || !data || data.length === 0) missing.push('revenue_daily (mês)')
      else moRevenue = data.reduce((s, r) => s + (Number(r.revenue) || 0), 0)
    }
    {
      const { data, error } = await supabase
        .from('google_ads_campaign_daily')
        .select('date, cost_micros')
        .gte('date', monthStart)
        .lte('date', yesterdayStr)
      if (error || !data || data.length === 0) missing.push('google_ads_campaign_daily (mês)')
      else moAds = data.reduce((s, r) => s + (Number(r.cost_micros) || 0), 0) / 1_000_000
    }
    {
      const { data, error } = await supabase
        .from('bureau_daily')
        .select('date, custo_bureau')
        .gte('date', monthStart)
        .lte('date', yesterdayStr)
      if (error || !data || data.length === 0) missing.push('bureau_daily (mês)')
      else moBureau = data.reduce((s, r) => s + (Number(r.custo_bureau) || 0), 0)
    }
    const moCogsKnown  = moAds !== null || moBureau !== null
    const moCogs       = (moAds ?? 0) + (moBureau ?? 0)
    const moLiquido    = moRevenue !== null && moCogsKnown ? moRevenue * 0.92 - moCogs : null
    const moMargem     = moRevenue !== null && moCogsKnown && moRevenue > 0 ? (moRevenue - moCogs) / moRevenue : null

    // ── Montar mensagem ──────────────────────────────────────────────────────
    const dataStr = new Date(yesterdayStr + 'T00:00:00Z').toLocaleDateString('pt-BR', {
      day: '2-digit', month: '2-digit', timeZone: 'UTC',
    })

    const lines: string[] = []
    lines.push(`📊 <b>Resumo VF — ${dataStr}</b>`)
    lines.push('')
    lines.push(
      lucroBruto !== null
        ? `📈 Lucro: ${fmtBRL(lucroBruto)}${margem !== null ? ` (margem ${(margem * 100).toFixed(1)}%)` : ''} | <b>Líquido: ${fmtBRL(lucroLiquido!)}</b>`
        : '📈 Lucro: — | <b>Líquido: —</b>'
    )
    lines.push(
      revenue !== null
        ? `💰 Receita: ${fmtBRL(revenue)} (${paidOrders ?? 0} pedidos${ticketMedio !== null ? `, ticket ${fmtBRL(ticketMedio)}` : ''})`
        : '💰 Receita: —'
    )
    lines.push(
      upsellOrders !== null
        ? `🔁 Upsell: ${upsellOrders} (${(upsellRate ?? 0).toFixed(1)}%)`
        : '🔁 Upsell: —'
    )
    lines.push(`📉 Estornos: ${refundCount} (${fmtBRL(refundValue)})`)
    lines.push(
      `🧲 Ads: ${custoAds !== null ? fmtBRL(custoAds) : '—'} | Bureau: ${custoBureau !== null ? fmtBRL(custoBureau) : '—'}`
    )
    lines.push('')
    lines.push(cacReal !== null ? `🎯 CAC real: ${fmtBRL(cacReal)}` : '🎯 CAC real: —')
    lines.push(
      balanceBRL !== null
        ? `🏦 Saldo Ads: ${fmtBRL(balanceBRL)}${daysLeft !== null ? ` (~${daysLeft.toFixed(1)} dias)` : ''}`
        : '🏦 Saldo Ads: —'
    )
    lines.push(
      dispatchTotal !== null
        ? `📟 Tracking: ${dispatchTotal} dispatches, ${dispatchFailed} falhas`
        : '📟 Tracking: —'
    )
    lines.push('____________________________________')
    lines.push('')
    lines.push('<b>Resultados deste Mês:</b>')
    lines.push(`Lucro líquido do mês: ${moLiquido !== null ? `<b>${fmtBRL(moLiquido)}</b>` : '—'}`)
    lines.push(`Receita do mês: ${moRevenue !== null ? fmtBRL(moRevenue) : '—'}`)
    lines.push(`COGS do mês: ${moCogsKnown ? fmtBRL(moCogs) : '—'}`)
    lines.push(`Margem do mês: ${moMargem !== null ? `${(moMargem * 100).toFixed(1)}%` : '—'}`)

    if (missing.length) {
      lines.push('')
      lines.push(`⚠️ sem dados: ${missing.join(', ')}`)
    }

    await sendTelegram(botToken!, chatId!, lines.join('\n'))

    return json({
      ok: true,
      date: yesterdayStr,
      revenue, paid_orders: paidOrders,
      upsell_orders: upsellOrders, upsell_rate: upsellRate,
      refund_count: refundCount, refund_value: refundValue,
      custo_ads: custoAds, custo_bureau: custoBureau,
      lucro_bruto: lucroBruto, lucro_liquido: lucroLiquido, margem,
      cac_real: cacReal,
      balance_brl: balanceBRL, days_left: daysLeft,
      dispatch_total: dispatchTotal, dispatch_failed: dispatchFailed,
      mes: { receita: moRevenue, cogs: moCogsKnown ? moCogs : null, lucro_liquido: moLiquido, margem: moMargem },
      missing,
    })
  } catch (err) {
    return json({ error: String(err) }, 500)
  }
})

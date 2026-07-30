/**
 * check-ads-upload-summary — watchdog diário dos diagnósticos de upload do Google Ads
 *
 * Consulta os dois recursos de diagnóstico de upload offline da API do Google Ads
 * e avisa no Telegram SÓ quando há notícia. Silêncio = tudo normal.
 *
 * Contexto (30/07/2026): os 3 alertas "Needs attention" do painel do Ads estavam
 * todos na fonte API, na ação `purchase`. Foi verificado que:
 *   - o envio está 100% (196 dispatches em 30/07, todos ads_status='success');
 *   - a aceitação no Google está EXCELLENT, sem alerts;
 *   - a agregação do Google roda ~02:30 UTC e cobre o DIA ANTERIOR, e o recurso
 *     por conversion action lag ~1 dia a mais que o de conta.
 * Ou seja, o painel em 30/07 ainda mostrava dado de antes da correção do order ID
 * (28/07 18:15 UTC). Esta função existe pra avisar quando isso deixa de ser verdade.
 *
 * Checagens:
 *   (a) failedCount > 0 em qualquer bucket recente (nível conta). Dedup por bucket,
 *       então cada dia problemático avisa uma vez só.
 *       ⚠️ É o sinal que valida o patch de 30/07 (adjustmentDateTime + userAgent):
 *       se o formato novo for rejeitado, aparece aqui.
 *   (b) status != EXCELLENT, ou campo `alerts` não vazio (nível conta).
 *   (c) agregação parada: nenhum bucket novo há mais de STALE_HOURS.
 *   (d) MARCO: lastUploadDateTime do nível conversion action alcançou 2026-07-30,
 *       que é o bucket do primeiro dia inteiramente pós-correção (29/07). A partir
 *       daí o painel da UI passa a refletir dado pós-correção e o coverage volta a
 *       significar alguma coisa. Avisa uma vez e nunca mais.
 *
 * Secrets necessários (todos já existem no projeto):
 *   GADS_CLIENT_ID / GADS_CLIENT_SECRET / GADS_REFRESH_TOKEN
 *   GADS_DEVELOPER_TOKEN / GADS_CUSTOMER_ID / GADS_PURCHASE_CONVERSION_ACTION_ID
 *   TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID
 *   SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY (injetados automaticamente)
 *
 * Deploy: supabase functions deploy check-ads-upload-summary --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:   cron-job.org → POST .../functions/v1/check-ads-upload-summary, 1x/dia às 03:00 UTC
 *         (a agregação do Google fecha ~02:30-02:50 UTC, então 03:00 pega fresquinho)
 *         Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { getGadsAccessToken } from '../_shared/conversion-dispatch.ts'

// Bucket carimbado no dia N cobre os uploads do dia N-1. A correção do order ID
// subiu em 28/07 18:15 UTC, então 29/07 é o primeiro dia inteiramente pós-correção
// e o bucket dele é o de 30/07. Comparação por STRING de data (YYYY-MM-DD) de
// propósito: o Google devolve 'YYYY-MM-DD HH:MM:SS.ffffff' sem fuso, e não está
// confirmado se é UTC ou o fuso da conta — comparar por dia evita depender disso.
const MARCO_DATA_UI_CONFIAVEL = '2026-07-30'
const STALE_HOURS = 48
const BUCKETS_RECENTES = 3
const DEDUP_WINDOW_HOURS = 20 * 24

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { 'Content-Type': 'application/json' } })
}

async function sendTelegram(botToken: string, chatId: string, text: string): Promise<void> {
  const resp = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: 'HTML' }),
  })
  if (!resp.ok) console.error('Telegram error:', await resp.text())
}

interface DailySummary { uploadDate?: string; successfulCount?: string; failedCount?: string; pendingCount?: string }
interface Summary {
  status?: string
  alerts?: unknown[]
  lastUploadDateTime?: string
  totalEventCount?: string
  successfulEventCount?: string
  pendingEventCount?: string
  dailySummaries?: DailySummary[]
}

// 'YYYY-MM-DD HH:MM:SS.ffffff' -> Date tratando como UTC. Só usado pra medir
// frescor com folga de 48h, onde um erro de 3h de fuso não muda a conclusão.
function parseAsUtc(s: string): Date | null {
  const m = s.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})/)
  return m ? new Date(`${m[1]}T${m[2]}Z`) : null
}

async function gaql(customerId: string, devToken: string, accessToken: string, query: string): Promise<Summary[]> {
  const resp = await fetch(`https://googleads.googleapis.com/v24/customers/${customerId}/googleAds:search`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'developer-token': devToken, 'Content-Type': 'application/json' },
    body: JSON.stringify({ query }),
  })
  const text = await resp.text()
  if (!resp.ok) throw new Error(`Google Ads API HTTP ${resp.status}: ${text}`)
  const parsed = JSON.parse(text) as { results?: Record<string, Summary>[] }
  // A chave do objeto varia com o recurso; pega o primeiro valor de cada result.
  return (parsed.results ?? []).map(r => Object.values(r)[0]).filter(Boolean)
}

Deno.serve(async (_req) => {
  const supabase = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!)

  const devToken   = Deno.env.get('GADS_DEVELOPER_TOKEN')
  const customerId = Deno.env.get('GADS_CUSTOMER_ID')
  const actionId   = Deno.env.get('GADS_PURCHASE_CONVERSION_ACTION_ID')
  const botToken   = Deno.env.get('TELEGRAM_BOT_TOKEN')
  const chatId     = Deno.env.get('TELEGRAM_CHAT_ID')
  const faltando = [
    !devToken   && 'GADS_DEVELOPER_TOKEN',
    !customerId && 'GADS_CUSTOMER_ID',
    !actionId   && 'GADS_PURCHASE_CONVERSION_ACTION_ID',
    !botToken   && 'TELEGRAM_BOT_TOKEN',
    !chatId     && 'TELEGRAM_CHAT_ID',
  ].filter(Boolean)
  if (faltando.length) return json({ ok: false, error: `secrets ausentes: ${faltando.join(', ')}` }, 500)

  const problemas: { key: string; line: string }[] = []

  try {
    const accessToken = await getGadsAccessToken()

    const C = 'offline_conversion_upload_client_summary'
    const A = 'offline_conversion_upload_conversion_action_summary'
    const [conta, porAcao] = await Promise.all([
      gaql(customerId!, devToken!, accessToken,
        `SELECT ${C}.client, ${C}.status, ${C}.alerts, ${C}.total_event_count, ${C}.successful_event_count, ` +
        `${C}.pending_event_count, ${C}.success_rate, ${C}.pending_rate, ${C}.last_upload_date_time, ${C}.daily_summaries ` +
        `FROM ${C}`),
      gaql(customerId!, devToken!, accessToken,
        `SELECT ${A}.conversion_action_id, ${A}.conversion_action_name, ${A}.status, ${A}.alerts, ` +
        `${A}.total_event_count, ${A}.successful_event_count, ${A}.last_upload_date_time, ${A}.daily_summaries ` +
        `FROM ${A} WHERE ${A}.conversion_action_id = ${actionId}`),
    ])

    const resumoConta = conta[0]
    const resumoAcao  = porAcao[0]
    if (!resumoConta) {
      problemas.push({ key: 'ads_upload_sem_resumo_conta', line: '⚠️ A API não devolveu resumo de upload no nível de conta.' })
    }

    // (a) buckets com falha
    const buckets = (resumoConta?.dailySummaries ?? []).slice(-BUCKETS_RECENTES)
    for (const b of buckets) {
      const failed = Number(b.failedCount ?? 0)
      if (failed > 0) {
        const dia = (b.uploadDate ?? '').slice(0, 10)
        problemas.push({
          key: `ads_upload_failed_${dia}`,
          line: `❌ Bucket <b>${dia}</b>: ${failed} upload(s) recusado(s) pelo Google (${b.successfulCount ?? '?'} aceitos).\n` +
                `   Se for o primeiro bucket depois de 30/07, suspeitar do patch de adjustmentDateTime/userAgent.`,
        })
      }
    }

    // (b) status / alerts
    if (resumoConta?.status && resumoConta.status !== 'EXCELLENT') {
      problemas.push({ key: `ads_upload_status_${resumoConta.status}`, line: `⚠️ Status do upload caiu para <b>${resumoConta.status}</b> (era EXCELLENT).` })
    }
    if (Array.isArray(resumoConta?.alerts) && resumoConta!.alerts!.length > 0) {
      problemas.push({ key: 'ads_upload_alerts', line: `⚠️ A API passou a reportar alerts: <code>${JSON.stringify(resumoConta!.alerts).slice(0, 500)}</code>` })
    }

    // (c) agregação parada
    const ultimoConta = resumoConta?.lastUploadDateTime ? parseAsUtc(resumoConta.lastUploadDateTime) : null
    if (ultimoConta) {
      const horas = (Date.now() - ultimoConta.getTime()) / 3600_000
      if (horas > STALE_HOURS) {
        problemas.push({
          key: `ads_upload_parado_${resumoConta!.lastUploadDateTime!.slice(0, 10)}`,
          line: `⚠️ Nenhum bucket novo há ${Math.round(horas)}h (último: ${resumoConta!.lastUploadDateTime}).\n` +
                `   Conferir se dispatch-conversions segue gravando ads_status='success'.`,
        })
      }
    }

    // (d) MARCO — painel da UI finalmente confiável
    const diaAcao = (resumoAcao?.lastUploadDateTime ?? '').slice(0, 10)
    if (diaAcao && diaAcao >= MARCO_DATA_UI_CONFIAVEL) {
      problemas.push({
        key: 'ads_upload_marco_ui_confiavel',
        line: `✅ <b>Marco atingido.</b> O diagnóstico por conversion action chegou em ${diaAcao} ` +
              `(bucket do primeiro dia inteiramente pós-correção do order ID).\n` +
              `   Agora vale remedir o coverage de <b>purchase</b> no painel do Ads — ` +
              `antes disso o número misturava dado pré-28/07.`,
      })
    }

    // Dedup
    const since = new Date(Date.now() - DEDUP_WINDOW_HOURS * 3600_000).toISOString()
    const { data: recentes, error: recentesErr } = await supabase
      .from('health_alerts').select('alert_key').gte('sent_at', since)
    if (recentesErr) return json({ ok: false, error: `health_alerts select: ${recentesErr.message}` }, 500)
    const jaAvisados = new Set((recentes ?? []).map(r => r.alert_key))
    const novos = problemas.filter(p => !jaAvisados.has(p.key))

    if (novos.length > 0) {
      const linhas = ['📡 <b>Diagnóstico de upload — Google Ads</b>', '', ...novos.map(p => p.line)]
      linhas.push('', `Conta: ${resumoConta?.status ?? '?'} · último bucket ${resumoConta?.lastUploadDateTime ?? '?'}`)
      linhas.push(`Ação purchase: último bucket ${resumoAcao?.lastUploadDateTime ?? '?'}`)
      await sendTelegram(botToken!, chatId!, linhas.join('\n'))
      await supabase.from('health_alerts').insert(novos.map(p => ({ alert_key: p.key })))
    }

    return json({
      ok: true,
      avisou: novos.length > 0,
      problemas_novos: novos.map(p => p.key),
      problemas_suprimidos_por_dedup: problemas.filter(p => jaAvisados.has(p.key)).map(p => p.key),
      conta: {
        status: resumoConta?.status ?? null,
        last_upload: resumoConta?.lastUploadDateTime ?? null,
        buckets_recentes: buckets,
      },
      acao_purchase: {
        status: resumoAcao?.status ?? null,
        last_upload: resumoAcao?.lastUploadDateTime ?? null,
        marco_atingido: Boolean(diaAcao && diaAcao >= MARCO_DATA_UI_CONFIAVEL),
      },
    })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    // Falha da própria checagem também merece aviso — watchdog mudo é pior que nenhum.
    await sendTelegram(botToken!, chatId!, `📡 <b>check-ads-upload-summary falhou</b>\n<code>${msg.slice(0, 800)}</code>`)
    return json({ ok: false, error: msg }, 500)
  }
})

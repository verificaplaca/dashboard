/**
 * check-bureau-spend — Monitoramento de gasto bureau (Assertiva) → Telegram alerts
 *
 * Roda 1x por dia (recomendado: 8h da manhã via cron-job.org).
 * Busca os dados de custo bureau já gravados na tabela `bureau_daily` do Supabase
 * e dispara alertas no Telegram quando detecta anomalias:
 *
 *   • daily_spike   → gasto de ontem > 30% acima da média dos últimos 7 dias
 *   • monthly_spike → projeção do mês atual > 40% acima do total do mês anterior
 *   • monthly_ok    → notificação diária matinal com resumo (sem anomalia)
 *
 * Cada tipo de alerta é disparado no máximo 1x por dia (tabela bureau_spend_alerts).
 *
 * Variáveis de ambiente (já configuradas do sistema de saldo Google Ads):
 *   TELEGRAM_BOT_TOKEN         Token do bot
 *   TELEGRAM_CHAT_ID           chat_id do grupo (-1003947525598)
 *   SUPABASE_URL               (injetado automaticamente)
 *   SUPABASE_SERVICE_ROLE_KEY  (injetado automaticamente)
 *
 * Deploy:  supabase functions deploy check-bureau-spend
 * Cron:    cron-job.org → POST .../functions/v1/check-bureau-spend todo dia às 8h
 *          Expressão: 0 8 * * *
 */

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// ─── Thresholds ───────────────────────────────────────────────────────────────
const DAILY_SPIKE_PCT   = 30;  // % acima da média 7d para alertar spike diário
const MONTHLY_SPIKE_PCT = 40;  // % acima do mês anterior para alertar projeção

// ─── Helpers ─────────────────────────────────────────────────────────────────
function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function fmtBRL(v: number): string {
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
}

function fmtPct(v: number): string {
  return (v >= 0 ? "+" : "") + v.toFixed(1) + "%";
}

async function sendTelegram(botToken: string, chatId: string, text: string): Promise<void> {
  const resp = await fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: "HTML" }),
  });
  if (!resp.ok) console.error("Telegram error:", await resp.text());
}

// ─── Main ────────────────────────────────────────────────────────────────────
Deno.serve(async () => {
  try {
    // ── Env vars ──────────────────────────────────────────────────────────
    const botToken = Deno.env.get("TELEGRAM_BOT_TOKEN");
    const chatId   = Deno.env.get("TELEGRAM_CHAT_ID");
    const supaUrl  = Deno.env.get("SUPABASE_URL");
    const supaKey  = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

    if (!botToken || !chatId || !supaUrl || !supaKey) {
      return json({ error: "Missing env vars" }, 500);
    }

    const supabase = createClient(supaUrl, supaKey);
    const today    = new Date();
    const todayStr = today.toISOString().slice(0, 10);

    // ── Buscar últimos 60 dias de bureau_daily ─────────────────────────────
    const since60 = new Date(today.getTime() - 60 * 86400000).toISOString().slice(0, 10);
    const { data: rows, error: fetchErr } = await supabase
      .from("bureau_daily")
      .select("date, custo_bureau")
      .gte("date", since60)
      .order("date", { ascending: true });

    if (fetchErr) return json({ error: fetchErr.message }, 500);
    if (!rows || rows.length === 0) return json({ ok: true, message: "Sem dados de bureau_daily." });

    // ── Organizar dados ────────────────────────────────────────────────────
    const parseDate = (s: string) => s.slice(0, 10);
    const allRows   = rows.map(r => ({ date: parseDate(String(r.date)), cost: parseFloat(String(r.custo_bureau)) || 0 }));

    // Mês atual e anterior
    const currMonth = todayStr.slice(0, 7); // "2026-05"
    const prevDate  = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const prevMonth = `${prevDate.getFullYear()}-${String(prevDate.getMonth() + 1).padStart(2, "0")}`;

    const currRows  = allRows.filter(r => r.date.startsWith(currMonth));
    const prevRows  = allRows.filter(r => r.date.startsWith(prevMonth));

    // Gasto acumulado no mês atual
    const currTotal = currRows.reduce((s, r) => s + r.cost, 0);

    // Gasto total do mês anterior
    const prevTotal = prevRows.reduce((s, r) => s + r.cost, 0);

    // Projeção do mês atual (ritmo atual × dias do mês)
    const daysInMonth   = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
    const daysPassed    = currRows.length || 1;
    const dailyAvgCurr  = currTotal / daysPassed;
    const projection    = dailyAvgCurr * daysInMonth;

    // Média dos últimos 7 dias (excluindo hoje se ainda não fechou)
    const last7   = allRows.slice(-8, -1).filter(r => r.cost > 0); // exclui último (hoje em aberto)
    const avg7    = last7.length ? last7.reduce((s, r) => s + r.cost, 0) / last7.length : 0;
    const yesterday = allRows[allRows.length - 2]; // penúltimo = ontem fechado
    const yesterdayCost = yesterday?.cost || 0;

    // ── Verificar alertas já enviados hoje ────────────────────────────────
    const { data: sentToday } = await supabase
      .from("bureau_spend_alerts")
      .select("alert_type")
      .eq("alert_date", todayStr);

    const sentTypes = new Set((sentToday || []).map((r: { alert_type: string }) => r.alert_type));

    const alerts: string[] = [];

    // ── Detecção 1: spike diário ───────────────────────────────────────────
    if (!sentTypes.has("daily_spike") && avg7 > 0 && yesterdayCost > 0) {
      const spikePct = ((yesterdayCost - avg7) / avg7) * 100;
      if (spikePct >= DAILY_SPIKE_PCT) {
        const msg = [
          `📈 <b>Spike de Gasto — Bureau Assertiva</b>`,
          ``,
          `Gasto de ontem está <b>${fmtPct(spikePct)}</b> acima da média de 7 dias.`,
          ``,
          `📊 Ontem: <b>${fmtBRL(yesterdayCost)}</b>`,
          `📊 Média 7d: <b>${fmtBRL(avg7)}</b>`,
          ``,
          `Verifique se houve aumento de volume de consultas ou mudança de produto.`,
        ].join("\n");

        await sendTelegram(botToken, chatId, msg);
        await supabase.from("bureau_spend_alerts").insert({
          alert_date: todayStr,
          alert_type: "daily_spike",
          details: { yesterdayCost, avg7, spikePct },
        });
        alerts.push("daily_spike");
      }
    }

    // ── Detecção 2: projeção mensal spike ──────────────────────────────────
    if (!sentTypes.has("monthly_spike") && prevTotal > 0) {
      const projPct = ((projection - prevTotal) / prevTotal) * 100;
      if (projPct >= MONTHLY_SPIKE_PCT) {
        const msg = [
          `📅 <b>Projeção Bureau acima do normal</b>`,
          ``,
          `Ritmo atual projeta gasto <b>${fmtPct(projPct)}</b> acima do mês anterior.`,
          ``,
          `📊 Projeção ${currMonth}: <b>${fmtBRL(projection)}</b>`,
          `📊 Total ${prevMonth}: <b>${fmtBRL(prevTotal)}</b>`,
          `📊 Acumulado agora: <b>${fmtBRL(currTotal)}</b> (${daysPassed}/${daysInMonth} dias)`,
          ``,
          `Verifique se o crescimento de consultas está acompanhado de receita proporcional.`,
        ].join("\n");

        await sendTelegram(botToken, chatId, msg);
        await supabase.from("bureau_spend_alerts").insert({
          alert_date: todayStr,
          alert_type: "monthly_spike",
          details: { projection, prevTotal, projPct, currTotal, daysPassed },
        });
        alerts.push("monthly_spike");
      }
    }

    // ── Resumo matinal diário (se não enviou nenhum alerta crítico) ────────
    if (!sentTypes.has("monthly_ok") && alerts.length === 0) {
      const projVsPrev = prevTotal > 0 ? ((projection - prevTotal) / prevTotal) * 100 : null;
      const projLine   = projVsPrev !== null
        ? `📊 Proj. mês: <b>${fmtBRL(projection)}</b> (${fmtPct(projVsPrev)} vs mês ant.)`
        : `📊 Proj. mês: <b>${fmtBRL(projection)}</b>`;

      const msg = [
        `🧾 <b>Resumo Bureau — ${new Date().toLocaleDateString("pt-BR")}</b>`,
        ``,
        `💰 Acumulado ${currMonth}: <b>${fmtBRL(currTotal)}</b>`,
        projLine,
        `📆 Média diária (7d): <b>${fmtBRL(avg7)}</b>`,
        `📆 Gasto ontem: <b>${fmtBRL(yesterdayCost)}</b>`,
        ``,
        `✅ Sem anomalias detectadas.`,
      ].join("\n");

      await sendTelegram(botToken, chatId, msg);
      await supabase.from("bureau_spend_alerts").insert({
        alert_date: todayStr,
        alert_type: "monthly_ok",
        details: { currTotal, projection, avg7, yesterdayCost },
      });
      alerts.push("monthly_ok");
    }

    return json({
      ok: true,
      date: todayStr,
      curr_total: currTotal,
      prev_total: prevTotal,
      projection,
      avg7,
      yesterday_cost: yesterdayCost,
      alerts_sent: alerts,
    });

  } catch (err) {
    return json({ error: String(err) }, 500);
  }
});

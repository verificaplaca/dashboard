/**
 * check-ads-balance — Google Ads prepaid balance → Supabase + Telegram alerts
 *
 * Busca o saldo disponível na conta Google Ads via REST API (GAQL),
 * salva o snapshot em `ads_balance_history` e dispara alerta no Telegram
 * quando o saldo cruza os thresholds: R$1.500 / R$1.000 / R$500.
 *
 * A notificação só é enviada UMA VEZ por threshold — o sistema verifica
 * o último nível notificado antes de disparar, evitando spam.
 *
 * Variáveis de ambiente necessárias (Supabase → Project Settings → Edge Functions):
 *   GADS_DEVELOPER_TOKEN       Developer Token do Google Ads
 *   GADS_CLIENT_ID             OAuth2 Client ID (Google Cloud Console)
 *   GADS_CLIENT_SECRET         OAuth2 Client Secret
 *   GADS_REFRESH_TOKEN         Refresh Token gerado via OAuth2 flow
 *   GADS_CUSTOMER_ID           ID da conta Google Ads (só números, sem traços)
 *   TELEGRAM_BOT_TOKEN         Token do bot (@BotFather)
 *   TELEGRAM_CHAT_ID           chat_id onde enviar as mensagens
 *   SUPABASE_URL               (já injetado automaticamente pelo Supabase)
 *   SUPABASE_SERVICE_ROLE_KEY  (já injetado automaticamente pelo Supabase)
 *
 * Deploy:  supabase functions deploy check-ads-balance
 * Cron:    cron-job.org → POST .../functions/v1/check-ads-balance a cada 1h
 *          Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

// ─── Thresholds em centavos (R$) ────────────────────────────────────────────
const THRESHOLDS = [
  { level: 1500, label: "⚠️ Aviso",    emoji: "🟡" },
  { level: 1000, label: "🔶 Atenção",  emoji: "🟠" },
  { level:  500, label: "🚨 Crítico",  emoji: "🔴" },
];

// ─── Helpers ─────────────────────────────────────────────────────────────────
function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function fmtBRL(value: number): string {
  return value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

// ─── Google Ads: obter access_token via refresh_token ────────────────────────
async function getAccessToken(
  clientId: string,
  clientSecret: string,
  refreshToken: string,
): Promise<string> {
  const resp = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id:     clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type:    "refresh_token",
    }),
  });
  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`OAuth2 token error: ${err}`);
  }
  const data = await resp.json();
  return data.access_token as string;
}

// ─── Google Ads: buscar saldo disponível via GAQL ────────────────────────────
// O saldo de contas pré-pagas fica em account_budget.
// balance = adjusted_spending_limit_micros − amount_served_micros
async function fetchBalance(
  accessToken: string,
  developerToken: string,
  customerId: string,
): Promise<number> {
  const apiVersion = "v17";
  const url = `https://googleads.googleapis.com/${apiVersion}/customers/${customerId}/googleAds:search`;

  const query = `
    SELECT
      account_budget.adjusted_spending_limit_micros,
      account_budget.amount_served_micros,
      account_budget.status
    FROM account_budget
    WHERE account_budget.status = 'APPROVED'
    LIMIT 1
  `;

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Authorization":          `Bearer ${accessToken}`,
      "developer-token":        developerToken,
      "Content-Type":           "application/json",
    },
    body: JSON.stringify({ query }),
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Google Ads API error ${resp.status}: ${err}`);
  }

  const data = await resp.json();
  const results = data.results ?? [];

  if (!results.length) {
    throw new Error("Nenhum AccountBudget aprovado encontrado para esta conta.");
  }

  const budget = results[0].accountBudget;
  const limit  = Number(budget.adjustedSpendingLimitMicros ?? 0);
  const served = Number(budget.amountServedMicros ?? 0);
  const balanceMicros = Math.max(0, limit - served);

  // Converte micros → R$ (1 BRL = 1_000_000 micros)
  return balanceMicros / 1_000_000;
}

// ─── Telegram: enviar mensagem ───────────────────────────────────────────────
async function sendTelegram(botToken: string, chatId: string, text: string): Promise<void> {
  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      chat_id:    chatId,
      text,
      parse_mode: "HTML",
    }),
  });
  if (!resp.ok) {
    const err = await resp.text();
    console.error("Telegram error:", err);
  }
}

// ─── Main ────────────────────────────────────────────────────────────────────
Deno.serve(async () => {
  try {
    // ── Env vars ──────────────────────────────────────────────────────────
    const developerToken = Deno.env.get("GADS_DEVELOPER_TOKEN");
    const clientId       = Deno.env.get("GADS_CLIENT_ID");
    const clientSecret   = Deno.env.get("GADS_CLIENT_SECRET");
    const refreshToken   = Deno.env.get("GADS_REFRESH_TOKEN");
    const customerId     = Deno.env.get("GADS_CUSTOMER_ID");
    const botToken       = Deno.env.get("TELEGRAM_BOT_TOKEN");
    const chatId         = Deno.env.get("TELEGRAM_CHAT_ID");
    const supaUrl        = Deno.env.get("SUPABASE_URL");
    const supaKey        = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

    const missing = [
      !developerToken && "GADS_DEVELOPER_TOKEN",
      !clientId       && "GADS_CLIENT_ID",
      !clientSecret   && "GADS_CLIENT_SECRET",
      !refreshToken   && "GADS_REFRESH_TOKEN",
      !customerId     && "GADS_CUSTOMER_ID",
      !botToken       && "TELEGRAM_BOT_TOKEN",
      !chatId         && "TELEGRAM_CHAT_ID",
      !supaUrl        && "SUPABASE_URL",
      !supaKey        && "SUPABASE_SERVICE_ROLE_KEY",
    ].filter(Boolean);

    if (missing.length) {
      return json({ error: `Missing env vars: ${missing.join(", ")}` }, 500);
    }

    const supabase = createClient(supaUrl!, supaKey!);

    // ── 1. Buscar saldo atual no Google Ads ───────────────────────────────
    const accessToken = await getAccessToken(clientId!, clientSecret!, refreshToken!);
    const balanceBRL  = await fetchBalance(accessToken, developerToken!, customerId!);

    // ── 2. Determinar threshold atingido ──────────────────────────────────
    let currentAlertLevel: number | null = null;
    for (const t of THRESHOLDS) {
      if (balanceBRL <= t.level) {
        currentAlertLevel = t.level;
        break; // pega o maior threshold que ainda é >= balance (ordem decrescente)
      }
    }

    // ── 3. Verificar último alerta enviado ────────────────────────────────
    const { data: lastRows } = await supabase
      .from("ads_balance_history")
      .select("alert_level")
      .order("checked_at", { ascending: false })
      .limit(1);

    const lastAlertLevel: number | null = lastRows?.[0]?.alert_level ?? null;

    // ── 4. Salvar snapshot ────────────────────────────────────────────────
    const { error: insertErr } = await supabase
      .from("ads_balance_history")
      .insert({
        balance_brl:  balanceBRL,
        alert_level:  currentAlertLevel,
        checked_at:   new Date().toISOString(),
      });

    if (insertErr) {
      return json({ error: `Supabase insert error: ${insertErr.message}` }, 500);
    }

    // ── 5. Enviar alerta Telegram (apenas quando threshold muda para pior) ─
    // Lógica: só envia se:
    //   a) cruzou um novo threshold (currentAlertLevel < lastAlertLevel ou era null)
    //   b) balance voltou ao normal: envia mensagem de "saldo ok"
    let alertSent = false;

    const levelChanged = currentAlertLevel !== lastAlertLevel;

    if (levelChanged && currentAlertLevel !== null) {
      // Cruzou um threshold — envia alerta
      const threshold = THRESHOLDS.find(t => t.level === currentAlertLevel)!;
      const msg = [
        `${threshold.emoji} <b>Alerta de Saldo Google Ads</b>`,
        ``,
        `${threshold.label}: saldo abaixo de <b>${fmtBRL(threshold.level)}</b>`,
        ``,
        `💰 Saldo atual: <b>${fmtBRL(balanceBRL)}</b>`,
        ``,
        `Recarregue sua conta para não interromper as campanhas.`,
      ].join("\n");

      await sendTelegram(botToken!, chatId!, msg);
      alertSent = true;

    } else if (levelChanged && currentAlertLevel === null && lastAlertLevel !== null) {
      // Saldo voltou ao normal após um alerta anterior
      const msg = [
        `✅ <b>Saldo Google Ads normalizado</b>`,
        ``,
        `💰 Saldo atual: <b>${fmtBRL(balanceBRL)}</b>`,
        ``,
        `Conta recarregada — campanhas funcionando normalmente.`,
      ].join("\n");

      await sendTelegram(botToken!, chatId!, msg);
      alertSent = true;
    }

    return json({
      ok:               true,
      balance_brl:      balanceBRL,
      alert_level:      currentAlertLevel,
      alert_sent:       alertSent,
      checked_at:       new Date().toISOString(),
    });

  } catch (err) {
    return json({ error: String(err) }, 500);
  }
});

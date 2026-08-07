/**
 * check-ads-balance — Google Ads prepaid balance → Supabase + Telegram alerts
 *
 * Busca o saldo disponível na conta Google Ads via REST API (GAQL),
 * salva o snapshot em `ads_balance_history` e dispara alerta no Telegram
 * APENAS quando o SALDO cai para R$ 1.500 ou menos. Acima disso: silêncio.
 *
 * Política de envio: alerta na hora quando o saldo cruza o limite pra baixo
 * (acima → abaixo). Enquanto continuar abaixo, reenvia 3x por dia (a cada 8h
 * desde o último envio real, rastreado via `alert_sent` em
 * `ads_balance_history`). Ao voltar acima do limite, envia aviso de recarga.
 *
 * O burn rate (média 7d de gasto real em `google_ads_campaign_daily`) NÃO
 * dispara alerta — entra na mensagem só como contexto (dias restantes e data
 * estimada de esgotamento).
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

// ─── Threshold único, baseado em VALOR do saldo ──────────────────────────────
// Alerta só existe abaixo (ou igual) desse valor. `alert_level` grava o próprio
// threshold (1500) quando o alerta está ativo, ou NULL quando o saldo está ok.
const BALANCE_THRESHOLD_BRL = 1500;

// Intervalo mínimo entre reenvios enquanto o saldo continua abaixo do limite.
// 8h => até 3 mensagens por dia (o cron roda de hora em hora).
const REALERT_INTERVAL_HOURS = 8;

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
  const apiVersion = "v23";
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

    // ── 2. Calcular burn rate (média 7d de gasto real) ────────────────────
    const today7 = new Date();
    const since7 = new Date(today7.getTime() - 8 * 86400000).toISOString().slice(0, 10);
    const { data: spendRows } = await supabase
      .from("google_ads_campaign_daily")
      .select("date, cost_micros")
      .gte("date", since7)
      .order("date", { ascending: false });

    // Agrupa por data e soma cost_micros
    const byDate = new Map<string, number>();
    for (const r of (spendRows ?? [])) {
      const d = String(r.date).slice(0, 10);
      byDate.set(d, (byDate.get(d) ?? 0) + Number(r.cost_micros ?? 0));
    }
    const dailyCosts = [...byDate.values()]
      .map(micros => micros / 1_000_000)
      .filter(v => v > 0)
      .slice(0, 7); // últimos 7 dias com gasto

    const avg7spend = dailyCosts.length
      ? dailyCosts.reduce((s, v) => s + v, 0) / dailyCosts.length
      : 0;

    const daysLeft = avg7spend > 0 ? balanceBRL / avg7spend : null;

    // ── 3. Determinar se o saldo está abaixo do limite ────────────────────
    // Regra única: saldo <= R$ 1.500 → alerta. Acima disso → nada.
    // Os dias restantes NÃO entram na decisão, só na mensagem.
    const belowThreshold = balanceBRL <= BALANCE_THRESHOLD_BRL;
    const currentAlertLevel: number | null = belowThreshold ? BALANCE_THRESHOLD_BRL : null;

    // ── 4. Verificar último nível e último envio real ──────────────────────
    const { data: lastRows } = await supabase
      .from("ads_balance_history")
      .select("alert_level")
      .order("checked_at", { ascending: false })
      .limit(1);

    const lastAlertLevel: number | null = lastRows?.[0]?.alert_level ?? null;

    const { data: lastSentRows } = await supabase
      .from("ads_balance_history")
      .select("checked_at")
      .eq("alert_sent", true)
      .order("checked_at", { ascending: false })
      .limit(1);

    const lastSentAt = lastSentRows?.[0]?.checked_at ? new Date(lastSentRows[0].checked_at) : null;
    const hoursSinceLastSent = lastSentAt ? (Date.now() - lastSentAt.getTime()) / 3600000 : null;

    // ── 5. Salvar snapshot ────────────────────────────────────────────────
    const { data: insertedRow, error: insertErr } = await supabase
      .from("ads_balance_history")
      .insert({
        balance_brl:  balanceBRL,
        alert_level:  currentAlertLevel,
        checked_at:   new Date().toISOString(),
      })
      .select("id")
      .single();

    if (insertErr) {
      return json({ error: `Supabase insert error: ${insertErr.message}` }, 500);
    }

    // ── 6. Enviar alerta Telegram ─────────────────────────────────────────
    // Cruzou o limite pra baixo (ok → abaixo de 1.500): envia na hora.
    // Continua abaixo: reenvia a cada 8h (3x/dia) desde o último envio real.
    // Voltou acima do limite (abaixo → ok): envia aviso de recarga.
    const crossedDown = currentAlertLevel !== null && lastAlertLevel === null;
    const recharged   = currentAlertLevel === null && lastAlertLevel !== null;
    const stillBelow  = currentAlertLevel !== null && lastAlertLevel !== null;
    const reAlert     = stillBelow && (
      lastSentAt === null ||
      (hoursSinceLastSent !== null && hoursSinceLastSent >= REALERT_INTERVAL_HOURS)
    );

    let alertSent = false;

    if (crossedDown || reAlert) {
      const runoutStr = daysLeft !== null
        ? new Date(Date.now() + daysLeft * 86400000)
            .toLocaleDateString("pt-BR", { day: "2-digit", month: "2-digit" })
        : null;

      const msg = [
        `🔴 <b>Alerta de Saldo Google Ads</b>`,
        ``,
        `💰 Saldo atual: <b>${fmtBRL(balanceBRL)}</b> (limite: ${fmtBRL(BALANCE_THRESHOLD_BRL)})`,
        avg7spend > 0 ? `📊 Gasto médio/dia (7d): <b>${fmtBRL(avg7spend)}</b>` : "",
        daysLeft !== null ? `⏳ Dias restantes: <b>${daysLeft.toFixed(1)} dias</b>` : "",
        runoutStr ? `📅 Esgotamento estimado: <b>${runoutStr}</b>` : "",
        ``,
        `Recarregue para não interromper as campanhas.`,
      ].filter(Boolean).join("\n");

      await sendTelegram(botToken!, chatId!, msg);
      alertSent = true;

    } else if (recharged) {
      // Saldo recarregado — voltou acima do limite
      const msg = [
        `✅ <b>Saldo Google Ads recarregado</b>`,
        ``,
        `💰 Saldo atual: <b>${fmtBRL(balanceBRL)}</b>`,
        daysLeft !== null ? `📅 Dias restantes: <b>${daysLeft.toFixed(1)} dias</b>` : "",
        ``,
        `Campanhas funcionando normalmente.`,
      ].filter(Boolean).join("\n");

      await sendTelegram(botToken!, chatId!, msg);
      alertSent = true;
    }

    if (alertSent && insertedRow?.id) {
      await supabase
        .from("ads_balance_history")
        .update({ alert_sent: true })
        .eq("id", insertedRow.id);
    }

    return json({
      ok:            true,
      balance_brl:   balanceBRL,
      threshold_brl: BALANCE_THRESHOLD_BRL,
      below_threshold: belowThreshold,
      days_left:     daysLeft,
      avg7_spend:    avg7spend,
      alert_level:   currentAlertLevel,
      alert_sent:    alertSent,
      checked_at:   new Date().toISOString(),
    });

  } catch (err) {
    return json({ error: String(err) }, 500);
  }
});

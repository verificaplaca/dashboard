/**
 * sync-refunds — Pagar.me /orders (status=refunded) → Supabase orders
 *
 * Busca pedidos com status "refunded" na Pagar.me dentro de uma janela
 * de lookback (padrão 72h) e faz upsert na tabela `orders`. A view
 * `refunds_daily` agrega automaticamente esses registros para o dashboard.
 *
 * Usa /orders (não /charges) pois reembolsos de PIX ficam no nível de pedido.
 *
 * Deploy:  supabase functions deploy sync-refunds
 * Cron:    cron-job.org → POST .../functions/v1/sync-refunds a cada 30min
 *          Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 *          Body: null
 */

import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const PAGARME_BASE   = "https://api.pagar.me/core/v5";
const PAGE_SIZE      = 100;
const MAX_PAGES      = 10;
const LOOKBACK_HOURS = 72; // pega estornos dos últimos 3 dias por execução

Deno.serve(async (_req) => {
  try {
    // ── Credenciais ────────────────────────────────────────────────────────
    const pagarmeKey = Deno.env.get("PAGARME_SECRET_KEY");
    const supaUrl    = Deno.env.get("SUPABASE_URL");
    const supaKey    = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

    if (!pagarmeKey || !supaUrl || !supaKey) {
      return json({ error: "Missing env vars: PAGARME_SECRET_KEY / SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY" }, 500);
    }

    const supabase = createClient(supaUrl, supaKey);

    // ── Janela de busca ────────────────────────────────────────────────────
    const now   = new Date();
    const since = new Date(now.getTime() - LOOKBACK_HOURS * 3600 * 1000);

    const params = new URLSearchParams({
      status:        "refunded",
      created_since: since.toISOString(),
      created_until: now.toISOString(),
      size:          String(PAGE_SIZE),
    });

    // ── Paginação Pagar.me /orders ─────────────────────────────────────────
    const orders: Record<string, unknown>[] = [];
    let   nextUrl: string | null = `${PAGARME_BASE}/orders?${params}`;
    let   page = 0;

    while (nextUrl && page < MAX_PAGES) {
      const resp = await fetch(nextUrl, {
        headers: { Authorization: "Basic " + btoa(pagarmeKey + ":") },
      });

      if (!resp.ok) {
        const body = await resp.text();
        return json({ error: `Pagar.me HTTP ${resp.status}`, detail: body }, 502);
      }

      const data = await resp.json();
      const items: Record<string, unknown>[] = Array.isArray(data.data) ? data.data : [];
      orders.push(...items);

      // Condição de parada: sem próxima página
      nextUrl = data.paging?.next ?? null;
      page++;
    }

    if (!orders.length) {
      return json({ ok: true, upserted: 0, message: "Nenhum estorno no período." });
    }

    // ── Mapeia para o schema de `orders` ───────────────────────────────────
    const rows = orders.map((o) => {
      const customer = (o.customer as Record<string, unknown>) ?? {};
      // amount em pedidos: campo `amount` (centavos)
      const amount = Number(o.amount ?? 0);
      return {
        provider:          "pagarme",
        provider_order_id: String(o.id   ?? ""),
        order_code:        String(o.code ?? ""),
        status:            "refunded",
        created_at:        o.created_at ?? null,
        updated_at:        o.updated_at ?? null,
        customer_id:       customer.id    ? String(customer.id)    : null,
        customer_email:    customer.email ? String(customer.email) : null,
        customer_name:     customer.name  ? String(customer.name)  : null,
        amount,
        currency:          String(o.currency ?? "BRL"),
        ingested_at:       now.toISOString(),
      };
    });

    // ── Upsert em `orders` ─────────────────────────────────────────────────
    const { error } = await supabase
      .from("orders")
      .upsert(rows, { onConflict: "provider_order_id" });

    if (error) return json({ error: error.message }, 500);

    // ── Log em sync_runs ───────────────────────────────────────────────────
    await supabase.from("sync_runs").insert({
      function_name:     "sync-refunds",
      records_processed: rows.length,
      started_at:        since.toISOString(),
      finished_at:       new Date().toISOString(),
    }).throwOnError().catch(() => {}); // não falha se a coluna não existir

    return json({ ok: true, upserted: rows.length });

  } catch (err) {
    return json({ error: String(err) }, 500);
  }
});

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

/**
 * sync-refunds — Pagar.me → Supabase orders (status = refunded)
 *
 * Busca cobranças com status "refunded" na Pagar.me dentro de uma janela
 * de lookback (padrão 72h) e faz upsert na tabela `orders`. A view
 * `refunds_daily` agrega automaticamente esses registros para o dashboard.
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

Deno.serve(async (req) => {
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
      status:       "refunded",
      created_since: since.toISOString(),
      created_until: now.toISOString(),
      size:          String(PAGE_SIZE),
    });

    // ── Paginação Pagar.me ─────────────────────────────────────────────────
    const charges: Record<string, unknown>[] = [];
    let   nextUrl: string | null = `${PAGARME_BASE}/charges?${params}`;
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
      charges.push(...items);

      // Condição de parada: sem próxima página
      nextUrl = data.paging?.next ?? null;
      page++;
    }

    if (!charges.length) {
      return json({ ok: true, upserted: 0, message: "Nenhum estorno no período." });
    }

    // ── Mapeia para o schema de `orders` ───────────────────────────────────
    const rows = charges.map((c) => {
      const customer = (c.customer as Record<string, unknown>) ?? {};
      return {
        provider:          "pagarme",
        provider_order_id: String(c.id   ?? ""),
        order_code:        String(c.code ?? ""),
        status:            "refunded",
        created_at:        c.created_at  ?? null,
        updated_at:        c.updated_at  ?? null,
        customer_id:       customer.id    ? String(customer.id)    : null,
        customer_email:    customer.email ? String(customer.email) : null,
        customer_name:     customer.name  ? String(customer.name)  : null,
        amount:            Number(c.amount ?? 0),   // centavos (bigint)
        currency:          String(c.currency ?? "BRL"),
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
      function_name: "sync-refunds",
      records_processed: rows.length,
      started_at:  since.toISOString(),
      finished_at: new Date().toISOString(),
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

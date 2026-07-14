-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_daily.sql
-- Função RPC get_bureau_costs_daily(p_start, p_end).
--
-- ⚠️ Roda no Supabase do SITE/SISTEMA, não no da DASHBOARD:
--    Site/Sistema (checkouts/bureau_chamadas_cobradas): https://ozquoloetuzynnyzkado.supabase.co
--    Dashboard (bureau_daily, dashboard.html):          https://ftmgmfdqdqxboiktxcoj.supabase.co
--
-- Cadeia real: cron-job.org (a cada 30min) → POST /functions/v1/sync-bureau-daily
-- (Edge Function na Dashboard) → chama esta RPC no Site/Sistema → upsert em
-- public.bureau_daily (Dashboard, onConflict: date).
--
-- ── HISTÓRICO ────────────────────────────────────────────────────────────
-- 2026-07-09: dev do site REESCREVEU esta função (sem avisar): o custo agora
--   vem de public.bureau_chamadas_cobradas(p_start, p_end) — chamadas de
--   bureau realmente cobradas — em vez da antiga tabela de tiers fixos por
--   plano/data (P1–P6 Assertiva + SKUs CheckTudo). Modelo novo é mais preciso.
--   A versão do dev veio SEM SECURITY DEFINER, o que quebrou a chamada via
--   PostgREST (rodava como anon → RLS de checkouts bloqueava tudo → retornava
--   [] sem erro). Dash ficou com bureau zerado 10–14/07.
-- 2026-07-14: arquitetura final acordada com o dev — a exposição pública era
--   real (RPC com definer + execute pra anon deixava qualquer um com a anon
--   key do site consultar receita/custo/margem). Resolução:
--   (a) REVOKE EXECUTE de public/anon/authenticated + GRANT só pra
--       service_role (statements no fim deste arquivo);
--   (b) secret BUREAU_SUPABASE_KEY (projeto do dashboard) trocado da anon key
--       pra SERVICE_ROLE key do site — bypassa RLS, então a função NÃO
--       precisa (e não deve) ter SECURITY DEFINER.
--
-- ⚠️ REGRA: se esta função for dropada/recriada, re-rodar os REVOKE/GRANT do
--   fim do arquivo (drop reseta as permissões pro default, que inclui PUBLIC).
-- ⚠️ A função public.bureau_chamadas_cobradas é mantida pelo dev do site e
--   NÃO está versionada neste repo.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION public.get_bureau_costs_daily(
  p_start date DEFAULT (date_trunc('month'::text, (now() AT TIME ZONE 'America/Sao_Paulo'::text)))::date,
  p_end   date DEFAULT ((now() AT TIME ZONE 'America/Sao_Paulo'::text))::date
)
RETURNS TABLE (
  dia           date,
  vendas_pagas  bigint,
  vendido_real  numeric,
  custo_bureau  numeric,
  lucro_bruto   numeric,
  margem_pct    numeric
)
LANGUAGE sql
STABLE
AS $function$
  WITH receita AS (
    SELECT
      DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo') AS dia,
      COUNT(*)::bigint  AS vendas_pagas,
      SUM(ck.valor)     AS vendido_real
    FROM checkouts ck
    WHERE (ck.paid_at IS NOT NULL OR LOWER(ck.status) = 'paid')
      AND LOWER(ck.status) <> 'refunded'
      AND ck.is_cortesia = false
      AND DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo')
        BETWEEN p_start AND p_end
      AND ck.plano NOT IN (
        'basico','premium','completo',
        'bin_estadual+bin_federal+gravame',
        'historico_leilao+indicio_sinistro'
      )
    GROUP BY 1
  ),
  custo AS (
    SELECT c.dia, SUM(c.custo) AS custo_bureau
    FROM public.bureau_chamadas_cobradas(p_start, p_end) c
    GROUP BY 1
  )
  SELECT
    COALESCE(r.dia, c.dia)                                    AS dia,
    COALESCE(r.vendas_pagas, 0)                               AS vendas_pagas,
    ROUND(COALESCE(r.vendido_real, 0)::numeric, 2)            AS vendido_real,
    ROUND(COALESCE(c.custo_bureau, 0)::numeric, 2)            AS custo_bureau,
    ROUND((COALESCE(r.vendido_real, 0) - COALESCE(c.custo_bureau, 0))::numeric, 2) AS lucro_bruto,
    ROUND(100.0 * (COALESCE(r.vendido_real, 0) - COALESCE(c.custo_bureau, 0))
      / NULLIF(r.vendido_real, 0), 1)                          AS margem_pct
  FROM receita r
  FULL JOIN custo c USING (dia)
  ORDER BY dia DESC;
$function$;

-- Permissões: só o pipeline (service_role) pode executar.
REVOKE EXECUTE ON FUNCTION public.get_bureau_costs_daily(date, date) FROM PUBLIC, anon, authenticated;
GRANT  EXECUTE ON FUNCTION public.get_bureau_costs_daily(date, date) TO service_role;

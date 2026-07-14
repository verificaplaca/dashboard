-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_by_type.sql
-- Função RPC get_bureau_costs_by_bureau(p_start, p_end).
--
-- ⚠️ Roda no Supabase do SITE/SISTEMA, não no da DASHBOARD:
--    Site/Sistema (checkouts/bureau_chamadas_cobradas): https://ozquoloetuzynnyzkado.supabase.co
--    Dashboard (bureau_by_type_daily):                  https://ftmgmfdqdqxboiktxcoj.supabase.co
--
-- Cadeia: cron-job.org → POST /functions/v1/sync-bureau-by-type (Dashboard)
-- → chama esta RPC no Site/Sistema → upsert em public.bureau_by_type_daily.
--
-- ── HISTÓRICO ────────────────────────────────────────────────────────────
-- 2026-07-09: dev do site REESCREVEU esta função (junto com
--   get_bureau_costs_daily, ver bureau_daily.sql): custo agora vem de
--   public.bureau_chamadas_cobradas(p_start, p_end) — chamadas realmente
--   cobradas — em vez da tabela de tiers fixos por plano. Veio SEM SECURITY
--   DEFINER → RLS bloqueava a chamada via API → sync gravou 0 records
--   silenciosamente de 09 a 14/07.
-- 2026-07-14: SECURITY DEFINER restaurado via ALTER FUNCTION + backfill.
--   Arquivo atualizado para refletir produção, JÁ COM security definer
--   no CREATE (re-rodar este script é seguro).
--
-- ⚠️ REGRA: qualquer alteração DEVE manter SECURITY DEFINER, senão o sync
--   volta a gravar 0 records silenciosamente.
-- ⚠️ public.bureau_chamadas_cobradas é mantida pelo dev do site, não
--   versionada neste repo.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION public.get_bureau_costs_by_bureau(
  p_start date DEFAULT (date_trunc('month'::text, (now() AT TIME ZONE 'America/Sao_Paulo'::text)))::date,
  p_end   date DEFAULT ((now() AT TIME ZONE 'America/Sao_Paulo'::text))::date
)
RETURNS TABLE (
  dia           date,
  bureau        text,
  vendas_pagas  bigint,
  vendido_real  numeric,
  custo_bureau  numeric
)
LANGUAGE sql
STABLE SECURITY DEFINER
AS $function$
  WITH receita AS (
    SELECT
      DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo') AS dia,
      COALESCE(ck.bureau, 'desconhecido') AS bureau,
      COUNT(*)::bigint AS vendas_pagas,
      SUM(ck.valor)    AS vendido_real
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
    GROUP BY 1, 2
  ),
  custo AS (
    SELECT c.dia, c.bureau, SUM(c.custo) AS custo_bureau
    FROM public.bureau_chamadas_cobradas(p_start, p_end) c
    GROUP BY 1, 2
  )
  SELECT
    COALESCE(r.dia, c.dia)                          AS dia,
    COALESCE(r.bureau, c.bureau)                    AS bureau,
    COALESCE(r.vendas_pagas, 0)                     AS vendas_pagas,
    ROUND(COALESCE(r.vendido_real, 0)::numeric, 2)  AS vendido_real,
    ROUND(COALESCE(c.custo_bureau, 0)::numeric, 2)  AS custo_bureau
  FROM receita r
  FULL JOIN custo c USING (dia, bureau)
  ORDER BY dia DESC, bureau;
$function$;

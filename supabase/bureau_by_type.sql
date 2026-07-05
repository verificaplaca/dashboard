-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_by_type.sql
-- Função NOVA get_bureau_costs_by_bureau(p_start, p_end) — ADITIVA, não altera
-- nem substitui a função existente get_bureau_costs_daily() (supabase/bureau_daily.sql).
--
-- Objetivo: quebrar o custo de bureau por tipo (Assertiva vs CheckTudo), pra
-- entender quanto cada um está custando enquanto os dois ficam ligados em
-- paralelo. Mesma lógica de cálculo da função principal, só que agrupando
-- também por `bureau` (não só por `dia`).
--
-- Roda no Supabase do SITE/SISTEMA (https://ozquoloetuzynnyzkado.supabase.co),
-- igual à get_bureau_costs_daily — é lá que ficam checkouts/consultas_placa.
--
-- Consumida pela Edge Function sync-bureau-by-type (Dashboard), que grava o
-- resultado na tabela bureau_by_type_daily (supabase/bureau_by_type_daily.sql).
-- ─────────────────────────────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION public.get_bureau_costs_by_bureau(
  p_start date DEFAULT ((CURRENT_DATE - '90 days'::interval))::date,
  p_end   date DEFAULT CURRENT_DATE
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

WITH vendas AS (
  SELECT
    DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo') AS dia,
    ck.plano,
    COALESCE(ck.bureau, 'assertiva') AS bureau,  -- checkouts antigos sem bureau preenchido = assertiva (era o único disponível)
    ck.id AS checkout_id,
    ck.valor,
    COALESCE(bool_or(cp.from_cache), false) AS from_cache
  FROM checkouts ck
  LEFT JOIN consultas_placa cp ON cp.checkout_id = ck.id
  WHERE (ck.paid_at IS NOT NULL OR LOWER(ck.status) = 'paid')
    AND LOWER(ck.status) <> 'refunded'
    AND ck.is_cortesia = false
    AND DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo') BETWEEN p_start AND p_end
    AND ck.plano NOT IN (
      'basico','premium','completo',
      'bin_estadual+bin_federal+gravame',
      'historico_leilao+indicio_sinistro'
    )
  GROUP BY 1, 2, 3, 4, 5
),
com_custo AS (
  SELECT
    v.dia,
    v.bureau,
    v.valor,
    v.from_cache,
    CASE
      WHEN v.from_cache THEN 0
      WHEN v.plano = 'leilao' THEN 15.10
      WHEN v.bureau = 'checktudo' THEN
        CASE v.plano
          WHEN 'padrao'             THEN 1.24
          WHEN 'bin_estadual'       THEN 3.07
          WHEN 'bin_estadual_owner' THEN 1.35
          WHEN 'bin_federal'        THEN 1.30
          WHEN 'renavam'            THEN 0.00
          WHEN 'gravame'            THEN 1.35
          WHEN 'indicio_sinistro'   THEN 1.85
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro'                    THEN 7.57
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 7.57
          ELSE 0
        END
      WHEN v.dia < '2026-02-18' THEN
        CASE v.plano
          WHEN 'padrao'            THEN 2.279
          WHEN 'bin_estadual'      THEN 4.660
          WHEN 'bin_federal'       THEN 5.939
          WHEN 'renavam'           THEN 5.939
          WHEN 'gravame'           THEN 5.939
          WHEN 'indicio_sinistro'  THEN 4.727
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 23.544
          ELSE 0
        END
      WHEN v.dia < '2026-03-26' THEN
        CASE v.plano
          WHEN 'padrao'            THEN 2.368
          WHEN 'bin_estadual'      THEN 4.841
          WHEN 'bin_federal'       THEN 6.168
          WHEN 'renavam'           THEN 6.168
          WHEN 'gravame'           THEN 6.168
          WHEN 'indicio_sinistro'  THEN 4.773
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 24.318
          ELSE 0
        END
      WHEN v.dia < '2026-04-20' THEN
        CASE v.plano
          WHEN 'padrao'            THEN 1.514
          WHEN 'bin_estadual'      THEN 4.841
          WHEN 'bin_estadual_owner'THEN 4.841
          WHEN 'bin_federal'       THEN 4.874
          WHEN 'renavam'           THEN 4.874
          WHEN 'gravame'           THEN 5.084
          WHEN 'indicio_sinistro'  THEN 4.636
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 20.949
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 20.949
          ELSE 0
        END
      WHEN v.dia < '2026-05-01' THEN
        CASE v.plano
          WHEN 'padrao'            THEN 1.175
          WHEN 'bin_estadual'      THEN 4.841
          WHEN 'bin_estadual_owner'THEN 4.841
          WHEN 'bin_federal'       THEN 4.841
          WHEN 'renavam'           THEN 4.841
          WHEN 'gravame'           THEN 4.509
          WHEN 'indicio_sinistro'  THEN 4.545
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 17.736
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 17.736
          ELSE 0
        END
      WHEN v.dia < '2026-05-30' THEN
        CASE v.plano
          WHEN 'padrao'            THEN 1.083
          WHEN 'bin_estadual'      THEN 4.730
          WHEN 'bin_estadual_owner'THEN 4.730
          WHEN 'bin_federal'       THEN 4.730
          WHEN 'renavam'           THEN 4.730
          WHEN 'gravame'           THEN 4.299
          WHEN 'indicio_sinistro'  THEN 4.500
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 19.342
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 19.342
          ELSE 0
        END
      ELSE
        CASE v.plano
          WHEN 'padrao'            THEN 1.00
          WHEN 'bin_estadual'      THEN 4.451
          WHEN 'bin_estadual_owner'THEN 4.451
          WHEN 'bin_federal'       THEN 4.451
          WHEN 'renavam'           THEN 4.451
          WHEN 'gravame'           THEN 4.022
          WHEN 'indicio_sinistro'  THEN 4.250
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro' THEN 18.174
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 18.174
          ELSE 0
        END
    END AS custo_bureau
  FROM vendas v
)
SELECT
  dia,
  bureau,
  COUNT(*)::bigint              AS vendas_pagas,
  ROUND(SUM(valor)::numeric, 2) AS vendido_real,
  ROUND(SUM(custo_bureau)::numeric, 3) AS custo_bureau
FROM com_custo
GROUP BY dia, bureau
ORDER BY dia DESC, bureau;

$function$;

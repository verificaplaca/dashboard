-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_daily_comparativo.sql
-- Script de comparação ONE-OFF: custo de bureau calculado pela função ANTIGA
-- (pré-atualização, só Assertiva P1–P4, sem CheckTudo) vs a NOVA
-- (get_bureau_costs_daily atual, com P1–P6 + CheckTudo).
--
-- Rode no SQL Editor do Site/Sistema (projeto ozquoloetuzynnyzkado).
-- Não mexe na função de produção — cria uma função temporária só para o teste
-- e derruba ela no final.
-- ─────────────────────────────────────────────────────────────────────────────

-- 1) Recria a lógica ANTIGA (capturada via pg_get_functiondef em 2026-07-05,
--    antes da atualização) sob outro nome, só para comparação:
CREATE OR REPLACE FUNCTION public._get_bureau_costs_daily_old(
  p_start date DEFAULT ((CURRENT_DATE - '90 days'::interval))::date,
  p_end   date DEFAULT CURRENT_DATE
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

WITH vendas AS (
  SELECT
    DATE(ck.created_at AT TIME ZONE 'America/Sao_Paulo') AS dia,
    ck.plano,
    ck.id AS checkout_id,
    ck.valor,
    COALESCE(bool_or(cp.from_cache), false) AS from_cache
  FROM checkouts ck
  LEFT JOIN consultas_placa cp ON cp.checkout_id = ck.id
  WHERE (ck.paid_at IS NOT NULL OR LOWER(ck.status) = 'paid')
    AND ck.is_cortesia = false
    AND DATE(ck.created_at AT TIME ZONE 'America/Sao_Paulo') BETWEEN p_start AND p_end
    AND ck.plano NOT IN (
      'basico','premium','completo',
      'bin_estadual+bin_federal+gravame',
      'historico_leilao+indicio_sinistro'
    )
  GROUP BY 1, 2, 3, 4
),
com_custo AS (
  SELECT
    v.dia,
    v.plano,
    v.valor,
    v.from_cache,
    CASE
      WHEN v.from_cache THEN 0
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
      ELSE
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
    END AS custo_bureau
  FROM vendas v
)
SELECT
  dia,
  COUNT(*)::bigint                                    AS vendas_pagas,
  ROUND(SUM(valor)::numeric, 2)                       AS vendido_real,
  ROUND(SUM(custo_bureau)::numeric, 3)                AS custo_bureau,
  ROUND((SUM(valor) - SUM(custo_bureau))::numeric, 2) AS lucro_bruto,
  ROUND(100.0 * (SUM(valor) - SUM(custo_bureau)) / NULLIF(SUM(valor), 0), 1) AS margem_pct
FROM com_custo
GROUP BY dia
ORDER BY dia DESC;

$function$;

-- 2) Comparação lado a lado — ajuste o período em p_start/p_end conforme
--    quiser (ex: desde o início do P4 até hoje, pra pegar P4/P5/P6 e CheckTudo):
SELECT
  COALESCE(novo.dia, antigo.dia)                                   AS dia,
  antigo.custo_bureau                                              AS custo_antigo,
  novo.custo_bureau                                                AS custo_novo,
  ROUND((novo.custo_bureau - antigo.custo_bureau)::numeric, 2)     AS diferenca,
  antigo.lucro_bruto                                               AS lucro_antigo,
  novo.lucro_bruto                                                 AS lucro_novo,
  ROUND((novo.lucro_bruto - antigo.lucro_bruto)::numeric, 2)       AS diferenca_lucro
FROM public.get_bureau_costs_daily('2026-04-20', CURRENT_DATE) novo
FULL OUTER JOIN public._get_bureau_costs_daily_old('2026-04-20', CURRENT_DATE) antigo
  ON antigo.dia = novo.dia
ORDER BY dia DESC;

-- 3) Limpeza — derruba a função temporária depois de conferir os resultados:
-- DROP FUNCTION IF EXISTS public._get_bureau_costs_daily_old(date, date);

-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_daily.sql
-- Função RPC get_bureau_costs_daily() — é isso que roda de verdade.
--
-- Cadeia real: Apps Script (pagarme-unified.js:1218 syncBureauFromSupabase,
-- disparado por trigger agendado) → POST /rest/v1/rpc/get_bureau_costs_daily
-- no Supabase → grava resultado na aba "BureauDaily" da planilha → syncDashboard
-- consolida em "Dashboard" → dashboard.html lê a tabela bureau_daily do Supabase
-- (alimentada a partir da mesma fonte) para os cards de custo/lucro bruto.
--
-- Calcula custo de bureau por venda (Assertiva por tiers de data + CheckTudo
-- por SKU fixo) e agrega por dia.
--
-- Rode este script no Supabase SQL Editor sempre que houver mudança de tier de
-- preço (Assertiva) ou novo bureau — ele recria a função (CREATE OR REPLACE),
-- não precisa dropar antes. Este arquivo é a fonte da verdade versionada da
-- função; antes só existia direto no banco, sem cópia no repo.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE OR REPLACE FUNCTION get_bureau_costs_daily()
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
AS $$
WITH vendas AS (
  SELECT
    DATE(COALESCE(ck.paid_at, ck.created_at) AT TIME ZONE 'America/Sao_Paulo') AS dia,  -- era created_at; agora paid_at (fallback p/ criação)
    ck.plano,
    ck.bureau,
    ck.id AS checkout_id,
    ck.valor,
    COALESCE(bool_or(cp.from_cache), false) AS from_cache
  FROM checkouts ck
  LEFT JOIN consultas_placa cp ON cp.checkout_id = ck.id
  WHERE (ck.paid_at IS NOT NULL OR LOWER(ck.status) = 'paid')
    AND LOWER(ck.status) <> 'refunded'          -- exclui estornos da receita
    AND ck.is_cortesia = false
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
    v.plano,
    v.bureau,
    v.valor,
    v.from_cache,
    CASE
      WHEN v.from_cache THEN 0
      -- Produto Leilão (q68) é SEMPRE CheckTudo, independe do bureau do checkout
      WHEN v.plano = 'leilao' THEN 15.10
      -- ═══════════ CheckTudo (plano 6k assinado) — custo por SKU ═══════════
      WHEN v.bureau = 'checktudo' THEN
        CASE v.plano
          WHEN 'padrao'             THEN 1.24   -- q1(0.21)+q3(1.03)
          WHEN 'bin_estadual'       THEN 3.07   -- q5888(1.72)+q4(1.35)
          WHEN 'bin_estadual_owner' THEN 1.35   -- q4
          WHEN 'bin_federal'        THEN 1.30   -- q11
          WHEN 'renavam'            THEN 0.00   -- vem na q3 (sem consulta nova)
          WHEN 'gravame'            THEN 1.35   -- q34
          WHEN 'indicio_sinistro'   THEN 1.85   -- q210
          WHEN 'bin_estadual+bin_federal+renavam+gravame+indicio_sinistro'                    THEN 7.57
          WHEN 'bin_estadual+bin_estadual_owner+bin_federal+renavam+gravame+indicio_sinistro' THEN 7.57
          ELSE 0
        END
      -- ═══════════ Assertiva — tiers por data (inalterado) ═══════════
      -- ── P1: início – 17/02/2026 ──
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
      -- ── P2: 18/02 – 25/03/2026 ──
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
      -- ── P3: 26/03 – 19/04/2026 ──
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
      -- ── P4: 20/04 – 30/04/2026 (Plano R$ 9.000) ──
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
      -- ── P5: 01/05 – 29/05/2026 (Plano R$ 14.000) ──
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
      -- ── P6: 30/05/2026+ (Plano R$ 18.000) ──
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
  COUNT(*)::bigint                                    AS vendas_pagas,
  ROUND(SUM(valor)::numeric, 2)                       AS vendido_real,
  ROUND(SUM(custo_bureau)::numeric, 3)                AS custo_bureau,
  ROUND((SUM(valor) - SUM(custo_bureau))::numeric, 2) AS lucro_bruto,
  ROUND(100.0 * (SUM(valor) - SUM(custo_bureau)) / NULLIF(SUM(valor), 0), 1) AS margem_pct
FROM com_custo
GROUP BY dia
ORDER BY dia DESC;
$$;

-- Breakdown por plano/bureau/dia (debug / auditoria manual — não faz parte da
-- função RPC, é só para rodar solto no SQL Editor quando precisar investigar):
/*
WITH vendas AS ( ... mesmo CTE acima ... ), com_custo AS ( ... mesmo CASE acima ... )
SELECT
  plano,
  bureau,
  dia,
  COUNT(*)::bigint                                          AS vendas_pagas,
  COUNT(*) FILTER (WHERE from_cache)::bigint                AS vendas_cache,
  ROUND(100.0 * COUNT(*) FILTER (WHERE from_cache) / NULLIF(COUNT(*), 0), 1) AS cache_hit_pct,
  ROUND(SUM(valor)::numeric, 2)                             AS vendido_real,
  ROUND(AVG(valor)::numeric, 2)                             AS ticket_medio,
  ROUND(SUM(custo_bureau)::numeric, 2)                      AS custo_bureau,
  ROUND(AVG(custo_bureau)::numeric, 3)                      AS custo_medio,
  ROUND((SUM(valor) - SUM(custo_bureau))::numeric, 2)       AS lucro_bruto,
  ROUND(100.0 * (SUM(valor) - SUM(custo_bureau)) / NULLIF(SUM(valor), 0), 1) AS margem_pct
FROM com_custo
GROUP BY plano, bureau, dia
ORDER BY lucro_bruto DESC;
*/

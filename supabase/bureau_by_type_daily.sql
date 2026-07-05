-- ─────────────────────────────────────────────────────────────────────────────
-- bureau_by_type_daily.sql
-- Tabela NOVA no Supabase da DASHBOARD (https://ftmgmfdqdqxboiktxcoj.supabase.co)
-- para guardar o custo de bureau quebrado por tipo (Assertiva vs CheckTudo).
--
-- Alimentada pela Edge Function sync-bureau-by-type, que chama
-- get_bureau_costs_by_bureau() no Site/Sistema (supabase/bureau_by_type.sql)
-- e faz upsert aqui. Não substitui nem altera a tabela bureau_daily existente
-- — é uma tabela nova, em paralelo.
-- ─────────────────────────────────────────────────────────────────────────────

CREATE TABLE IF NOT EXISTS public.bureau_by_type_daily (
  date          date        NOT NULL,
  bureau        text        NOT NULL,
  vendas_pagas  bigint      NOT NULL DEFAULT 0,
  vendido_real  numeric     NOT NULL DEFAULT 0,
  custo_bureau  numeric     NOT NULL DEFAULT 0,
  ingested_at   timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (date, bureau)
);

-- Leitura pública (mesmo padrão de bureau_daily, ajuste conforme suas policies
-- de RLS existentes se a tabela bureau_daily já usar uma role/policy específica).
ALTER TABLE public.bureau_by_type_daily ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "bureau_by_type_daily_select" ON public.bureau_by_type_daily;
CREATE POLICY "bureau_by_type_daily_select" ON public.bureau_by_type_daily
  FOR SELECT USING (true);

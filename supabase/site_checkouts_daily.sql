-- ─────────────────────────────────────────────────────────────────────────────
-- VIEW checkouts_daily — checkouts iniciados por dia (BRT)
-- Rodar no Supabase do SITE (ozquoloetuzynnyzkado), SQL Editor.
--
-- OBJETIVO: alimentar a Taxa de Conversão REAL do dashboard (compras ÷
-- checkouts iniciados). Hoje o dashboard não tem dados de checkout
-- (o card mostra "—"); numa etapa futura o dashboard.html vai consultar
-- esta view e calcular convRate = paid_orders ÷ checkouts_started.
-- A integração no dashboard fica para depois — este script só cria a view.
--
-- Agrupamento em America/Sao_Paulo para casar com bureau_daily/refunds_daily
-- e evitar que checkouts pós-21h BRT caiam no dia seguinte (UTC).
-- ─────────────────────────────────────────────────────────────────────────────

create or replace view checkouts_daily as
select
  date(created_at at time zone 'America/Sao_Paulo') as date,
  count(*)                                          as checkouts_started
from checkouts
group by 1
order by 1;

-- Leitura pública (o dashboard usa a anon key, somente SELECT)
grant select on checkouts_daily to anon;

-- Teste rápido:
-- select * from checkouts_daily order by date desc limit 14;

-- ============================================================================
-- VIEW checkouts_campaign_daily — checkouts PAGOS por dia × utm_campaign × plano
-- Rodar no Supabase do SITE (ozquoloetuzynnyzkado), SQL Editor —
-- NÃO no projeto do dashboard.
--
-- Cross-check de atribuição do painel "Por Produto — Teste Pacote Completo":
-- o dashboard compara as vendas do produto R$79,90 (view revenue_daily_completa,
-- Pagar.me/projeto do dashboard) com os checkouts pagos cujo utm_campaign é o
-- da campanha do teste. Se o % divergir muito de 100%, ou o utm da LP está
-- quebrado ou há tráfego de outra origem comprando o produto.
--
-- Pré-requisitos (já em produção desde 13-14/07/2026):
--   • colunas de atribuição em checkouts (site_add_attribution_columns.sql)
--   • attribution capture no site (P1.1) preenchendo utm_campaign
--
-- Convenções (iguais a checkouts_daily, mesmo projeto):
--   • dia em America/Sao_Paulo
--   • pago = paid_at IS NOT NULL (campo confiável; status='paid' tem anomalias
--     com paid_at nulo — ver decisão do tracking gateway 12/07)
--
-- Nota: linhas com utm_campaign vazio agregam em '(sem utm)' — inclui tráfego
-- orgânico/direto e checkouts anteriores ao attribution capture (13/07).
-- ============================================================================

create or replace view public.checkouts_campaign_daily as
select
  date(c.paid_at at time zone 'America/Sao_Paulo')  as date,
  coalesce(nullif(c.utm_campaign, ''), '(sem utm)') as utm_campaign,
  coalesce(nullif(c.plano, ''), '(sem plano)')      as plano,
  count(*)                                          as paid_count
from public.checkouts c
where c.paid_at is not null
group by 1, 2, 3;

-- Dashboard lê com a anon key do site (somente SELECT) — mesmo padrão de
-- checkouts_daily. Exposição: só contagens agregadas por dia/campanha/plano.
grant select on public.checkouts_campaign_daily to anon;

-- Teste rápido:
-- select * from checkouts_campaign_daily order by date desc, paid_count desc limit 30;
--
-- Quando a campanha do teste subir, conferir o valor EXATO do utm_campaign que
-- a LP está gravando (o dashboard casa por substring 'COMPLETA', case-insensitive):
-- select distinct utm_campaign from checkouts
-- where paid_at > now() - interval '7 days' and utm_campaign is not null;

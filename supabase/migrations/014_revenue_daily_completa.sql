-- ============================================================================
-- 014 — VIEW revenue_daily_completa
-- Rodar no Supabase do DASHBOARD (ftmgmfdqdqxboiktxcoj), SQL Editor.
--
-- Vendas diárias do PACOTE COMPLETO R$79,90 (teste jul/2026) a partir de
-- `orders` (Pagar.me). Alimenta o painel "Por Produto — Teste Pacote Completo"
-- do dashboard.html (série TESTE_DAILY: vendas/receita por dia).
--
-- Mesmas convenções das views analíticas existentes (revenue_daily etc.):
--   • data = COALESCE(paid_at, created_at) em America/Sao_Paulo
--   • status pagos = paid / delivered / authorized
--
-- ⚠ CRITÉRIO DO PRODUTO — PLACEHOLDER ATÉ A 1ª VENDA REAL:
--   amount = 7990 (R$79,90 em centavos) identifica o pedido do pacote hoje
--   porque nenhum outro pedido tem esse valor (principal = 1499; addons têm
--   outros valores). DOIS riscos conhecidos:
--     1. Cupom/desconto muda o amount e o pedido escaparia do filtro.
--     2. Se algum addon/combo futuro custar exatamente R$79,90, entraria aqui.
--   Assim que a 1ª venda real cair, validar com a query no fim deste arquivo e,
--   se o novo plano tiver item_code/descrição própria em order_items, trocar o
--   WHERE por um JOIN em order_items (critério por código > critério por valor).
-- ============================================================================

create or replace view public.revenue_daily_completa as
select
  date(coalesce(o.paid_at, o.created_at) at time zone 'America/Sao_Paulo') as date,
  count(*)                                                                 as vendas,
  round(sum(o.amount) / 100.0, 2)                                          as receita
from public.orders o
where lower(o.status) in ('paid', 'delivered', 'authorized')
  and o.amount = 7990
group by 1
order by 1;

-- Dashboard lê com a anon key (somente SELECT) — mesmo padrão de revenue_daily
grant select on public.revenue_daily_completa to anon;

-- ── Validação (rodar quando a 1ª venda real do pacote cair) ──────────────────
-- 1. O pedido apareceu na view?
--    select * from revenue_daily_completa order by date desc limit 7;
-- 2. Conferir o pedido cru e o que o identifica (amount? código? itens?):
--    select o.provider_order_id, o.order_code, o.amount, o.status, o.paid_at
--    from orders o
--    where o.amount = 7990
--    order by o.created_at desc limit 10;
-- 3. Ver itens do pedido pra descobrir o item_code do novo plano (se existir,
--    é o critério definitivo — atualizar o WHERE desta view):
--    select i.* from order_items i
--    join orders o on o.provider = i.provider
--                 and o.provider_order_id = i.provider_order_id
--    where o.amount = 7990
--    order by o.created_at desc limit 10;

-- Cover the addon join and aggregates without reading each item's heap page.
create index concurrently if not exists ix_order_items_analytics_addons
  on public.order_items (provider, provider_order_id)
  include (addon_key, total_amount)
  where addon_key is not null;

-- The complete-package graph only needs paid orders for this specific price.
create index concurrently if not exists ix_orders_analytics_complete_package
  on public.orders (((coalesce(paid_at, created_at) at time zone 'America/Sao_Paulo')::date))
  include (paid_at, created_at, amount)
  where amount = 7990 and lower(status) in ('paid', 'delivered', 'authorized');

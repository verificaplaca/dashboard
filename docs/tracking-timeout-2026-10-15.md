# Tracking — correção de indisponibilidade em 15/10/2026

O endpoint público revenue_daily retornava HTTP 500 / PostgreSQL 57014
(statement timeout). upsell_daily também excedia o limite. O projeto estava ativo,
e conversion_dispatches_public continuava recebendo envios confirmados.

EXPLAIN ANALYZE da consulta de receita dos últimos 30 dias mostrou seq scan em
orders (50.934 registros; tabela 119 MB), filtro por data calculada em São Paulo,
e execução de 13.811,497 ms. Os índices separados paid_at/created_at não atendiam
a expressão usada pelas views.

A migração 019 foi aplicada individualmente via `supabase db query --linked --file`,
sem db push, sem alteração de views, dados, RLS, cron, sync ou funções de envio.
Cria índice concorrente por data calculada, cobrindo as colunas analíticas.
O índice foi confirmado indisvalid/indisready. Plano posterior: Index Only Scan,
execução de 6,690 ms. As quatro consultas REST de receita/upsell (30 dias e janela
anual) voltaram a HTTP 200. Tempo HTTP, incluindo rede, entre 2,28 e 4,55 s na amostra.

Reconciliação antes/depois, na mesma fotografia sem novos pagamentos entre leituras:

| Fonte | Antes/depois |
|---|---|
| orders | 50.934 / 50.934 |
| order_items | 50.934 / 50.934 |
| revenue_daily — hash de todas as linhas ordenadas | c2bf2ddd1b7c62cee94f50dd34ac08e1 — idêntico |
| upsell_daily — hash de todas as linhas ordenadas | 3914d0a5f4d5df9f5261f5e5cd18a930 — idêntico |
| upsell_by_type — hash de todas as linhas ordenadas | f0ab49ea96bdfbe352a4ea2d47178404 — idêntico |

O frontend tracking-gateway.html carregava dispatches e CAC num único Promise.all:
a falha financeira descartava toda a apresentação do tracking. Agora as seções
carregam independentemente; falha no comparativo aparece apenas no CAC. Falha da
fonte de tracking limpa dados antigos, sem representar indisponibilidade como zero
ou saudável. O overlay libera o tracking assim que seus dados chegam. Erros mostram
código/mensagem e fonte corretos; 401/403/404 não sofrem cinco retentativas inúteis.
Atualizações antigas não sobrescrevem uma atualização mais nova.

Testes: `node --test tests/tracking-gateway.test.mjs`, cobrindo as falhas isoladas,
limpeza de dados antigos, concorrência, erro de autorização, escape do erro e
carregamento do tracking antes do CAC. Workflow tracking-checks executa em PR/push.
Console existente verificado com `npm --prefix console-app test`.

Validação visual com fonte real: 2.004 pedidos nos últimos sete dias, todos com
GA4/Google Ads/Meta success, zero pendências/falhas nessa janela. São status de envio
do gateway; o comparativo de atribuição CAC permanece separado (gap observado 6,7%).

Rollback do frontend: reverter somente os arquivos desta correção. O índice pode
permanecer; se necessário, DROP INDEX CONCURRENTLY public.ix_orders_analytics_payment_date
em comando isolado fora de transação, ciente de que a lentidão poderá retornar.
Nenhum registro deve ser apagado ou regravado para rollback.

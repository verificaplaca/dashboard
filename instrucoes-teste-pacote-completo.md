# Teste Pacote Completo R$79,90 — setup do painel e das views

O `dashboard.html` já está com o painel "Por Produto — Teste Pacote Completo"
(Visão Geral) e com a página Campanhas preparada para a campanha nova. Enquanto
as views abaixo não existirem, o painel mostra dados DEMO com chip amarelo.

## Régua do teste (benchmark 15/07/2026)

| Marco | CAC |
|---|---|
| Meta (vitória) | R$ 45–55 |
| Empate (mesma eficiência por real do funil base) | ~R$ 60 |
| Break-even | ~R$ 67 (bureau 1x; com bureau 3x cai pra ~R$ 62) |

## Passo 1 — View no Supabase do DASHBOARD (ftmgmfdqdqxboiktxcoj)

SQL Editor → colar e rodar `supabase/migrations/014_revenue_daily_completa.sql`.

Critério do produto: `amount = 7990` (centavos). É **placeholder até a 1ª venda
real** — cupom/desconto mudaria o amount e o pedido escaparia. O próprio arquivo
tem as queries de validação pra rodar quando a 1ª venda cair (e como migrar o
critério pra `order_items.item_code`, que é o definitivo).

## Passo 2 — View no Supabase do SITE (ozquoloetuzynnyzkado)

SQL Editor → colar e rodar `supabase/site_checkouts_campaign_daily.sql`.

Agrupa checkouts pagos (`paid_at is not null`) por dia × utm_campaign × plano.
É o cross-check de atribuição do painel.

## Passo 3 — Ao criar a campanha no Google Ads

1. **Nome da campanha deve conter `PACOTAO`** — campanha real criada em 16/07:
   `[SEARCH] [COMPRA] [LP-PACOTAO-1] [BEST] - JUL-2026`. É o `campMatch` no
   config `PRODUTO_TESTE` do dashboard.html — casa por substring,
   case-insensitive. Se renomear a campanha, manter "PACOTAO" no nome (ou
   ajustar `campMatch`).
2. **utm_campaign da LP também deve conter `pacotao`** (qualquer caixa) — o
   cross-check casa pelo mesmo `campMatch`. Conferir o valor exato gravado:
   query no fim de `site_checkouts_campaign_daily.sql`.
3. Ajustar `inicio` no config `PRODUTO_TESTE` pra data real de lançamento
   (hoje documentacional).
4. O custo entra sozinho: `sync-google-ads` já grava qualquer campanha em
   `google_ads_campaign_daily`.

## Como o painel se comporta

- **Views não existem ainda** → demo com chip "DADOS DEMO · views SQL pendentes".
- **Views criadas, campanha sem vendas** → painel real vazio ("sem dados no
  período"); sem demo. Dia com custo e 0 vendas aparece no gráfico (queima de
  verba visível).
- **Vendas reais chegando** → painel real; bloco "Consulta Base" passa a
  DESCONTAR o teste; na página Campanhas a linha `PACOTAO 🧪` usa vendas
  reais do produto (não rateio) e as demais campanhas rateiam o restante.

## Pendências fora do dashboard (bloqueiam dados, não código)

- Setup do funil no site: utm_campaign distinto + plan_id novo (79,90) no
  analytics.ts com plan_price correto; conferir eventos do funil curto.
- Confirmar custo real de bureau da consulta completa (config usa estimativa
  R$4,75/venda, 1x — a régua quase não muda até 3x).

## Commit (rodar no terminal, na pasta do repo)

```bash
git add dashboard.html supabase/migrations/014_revenue_daily_completa.sql supabase/site_checkouts_campaign_daily.sql instrucoes-teste-pacote-completo.md
git commit -m "Painel Por Produto (teste Pacote Completo R\$79,90) + campanhas dinâmicas + views do teste"
git pull --rebase origin main
git push origin main
```

# Verifica Placa — Contexto do Projeto
> Atualizar este arquivo sempre que houver mudanças estruturais no projeto.
> Usar como contexto inicial em novas conversas para economizar tokens.

---

## Produto
Dashboard de performance para o produto Verifica Placa.
Hospedado em GitHub Pages. Dados via Supabase (principal) + Google Sheets (legado/fallback).

---

## Repositórios e Pastas

| Pasta | Propósito |
|---|---|
| `/Users/danielmacedo/Documents/GitHub/verificaplaca/` | Repo do GitHub Pages — arquivos públicos |
| `/Users/danielmacedo/Documents/Claude/Projects/verificaplaca/` | Supabase, migrações SQL, edge functions |

**Git remote:** `github-verificaplaca:verificaplaca/dashboard.git`
**SSH config alias:** `github-verificaplaca` → `~/.ssh/id_ed25519_git_verificaplaca`
**Nunca usar `git push --force`. Conflito → `git pull --rebase origin main` depois push.**

---

## Arquivos Principais

| Arquivo | Descrição |
|---|---|
| `dashboard.html` | Dashboard principal — HTML + CSS + JS inline, arquivo único |
| `index.html` | Landing page |
| `CLAUDE.md` | Instruções de comportamento para Claude (git, SSH, convenções) |
| `supabase/functions/sync-orders-incremental/index.ts` | Sync Pagar.me → Supabase |
| `supabase/functions/sync-bureau-daily/index.ts` | Sync custo bureau → Supabase |
| `supabase/functions/sync-google-ads/index.ts` | Sync Google Ads → Supabase |
| `supabase/migrations/` | Migrações SQL em ordem numérica (001 → 006) |

---

## Supabase

**Projeto:** `ftmgmfdqdqxboiktxcoj.supabase.co`
**Plano:** Free

### Tabelas principais
- `orders` — pedidos do Pagar.me (chave: `provider, provider_order_id`)
- `order_items` — itens dos pedidos (chave: `provider, provider_order_id, provider_item_id`)
- `sync_runs` — histórico de execuções dos syncs
- `sync_errors` — erros dos syncs
- `google_ads_campaign_daily` — dados diários de campanhas
- `bureau_daily` — consultas de bureau por dia
- `monthly_targets` — metas mensais do sistema (1 linha/mês; RLS anon-read, auth-write)

### Views analíticas (usam `COALESCE(paid_at, created_at)` para agrupar por data)
- `revenue_daily` — receita e pedidos pagos por dia
- `upsell_daily` — taxa e volume de upsell por dia
- `upsell_by_type` — breakdown de upsell por addon_key

### Status válidos de pedido
`paid`, `delivered`, `authorized` (os três contam como receita)

---

## Sync de Dados — cron-job.org

Todos os crons foram migrados do GitHub Actions para cron-job.org (2026-05-01).
**Não recriar schedules no GitHub.** Workflows YML têm só `workflow_dispatch` (manual).

| Job | URL | Schedule (UTC) | Body |
|---|---|---|---|
| sync-orders-incremental | `.../functions/v1/sync-orders-incremental` | `0,30 * * * *` (a cada 30min) | `null` |
| sync-bureau-daily | `.../functions/v1/sync-bureau-daily` | min 0, horas 9/15/21 (= 6h/12h/18h BRT) | `{}` |
| sync-google-ads | `.../functions/v1/sync-google-ads` | min 30, horas 9/15/21 (= 6h30/12h30/18h30 BRT) | `{}` |
| retry-conversion-dispatches | `.../functions/v1/retry-conversion-dispatches` | a cada 15min | `{}` — pendente de criar, ver `instrucoes-setup-passos-5-7.md` |

**Headers em todos:** `Authorization: Bearer <VP_SERVICE_ROLE_KEY>` · `Content-Type: application/json`

---

## Dashboard — Convenções

### Cálculos principais
```js
// Lucro Bruto
pft = receita − custo_ads − custo_bureau

// Lucro Líquido — imposto FASEADO por data de venda (netFactor() no dashboard.html):
//   até mar/2026 → 0.92 | abr/2026 → 0.95 | mai/2026+ → 0.962
// O card da dash usa netPftOpt (fator por dia). NÃO usar 0.92 fixo (desatualizado).
netPftOpt = Σ dia: revenue(dia) * netFactor(dia) − costTotal(dia)

// EBITDA — sempre todo o período, NÃO muda com filtro de data
// (na prática é LUCRO BRUTO acumulado, não EBITDA contábil)
ALL_DAYS.reduce((s,d) => s + (d.profit||0), 0)
```

### Metas — tabela `monthly_targets` (Supabase principal)
As metas do sistema (CAC, upsell, orçamento Ads, receita, lucro bruto, lucro líquido)
vivem na tabela `monthly_targets` (1 linha por mês, `month` sempre dia 01), editáveis
no módulo **"Metas"** da sidebar do dashboard (migration `013_monthly_targets.sql`).

- **Leitura (anon):** `targetFor(dateStr)` no dashboard.html — cascata: linha do mês →
  linha anterior mais recente (meta vale até ser alterada) → constantes fallback do
  código (dash nunca quebra). Agregados do período usam a meta do mês do último dia
  do filtro (`activeTargets()`); linhas por-dia dos gráficos usam `targetFor(dia)`.
- **Escrita:** Supabase Auth email+senha (REST, sem SDK) — RLS só aceita
  INSERT/UPDATE de `authenticated`. Signup público desabilitado; usuário criado
  manualmente no painel.
- `google-ads-cac.html` lê a mesma tabela (`loadCacTarget()`, fallback `CAC_TARGET = 9`).
- As constantes `TARGET_*`/`MONTHLY_*` no topo do JS do dashboard são **só fallback**.
- Meta de CAC vigente (jul/26): **R$ 9** (o antigo TARGET_CPA=10 estava errado).

### Estrutura dos KPI Cards
- **Row 1 (5 cards):** Receita Bruta · Custo Total COGS · Lucro Bruto · Lucro Líquido · EBITDA
- **Row 2 (5 cards):** ROAS · CPV · CAC · Gasto Google Ads · Gasto Bureau
- **Row 3 (5 cards):** Margem · Taxa de Conversão · Taxa de Upsell · Transações · Ticket Médio

### Tooltips
- `has-tip` e `data-tip` ficam no `div.kpi-card` (não em spans internos)
- CSS: `top: 28px` (abaixo do ícone ⓘ), `z-index: 9999`
- `.kpi-card` precisa ter `overflow: visible` (para tooltip sair do card) e `min-width: 0` (evita overflow no grid mobile)

### Mobile (≤700px)
- Grid KPI: `repeat(2, 1fr)`
- 5º card (último em linha ímpar): `grid-column: 1/-1` para evitar lacuna vazia
- Estornos: classe `.refund-grid` → empilha em `1fr` no mobile
- Tabela upsell por tipo: coluna `.tendencia-col` oculta no mobile; seta de tendência aparece em `.trend-in-receita` junto à coluna Receita
- `kpi-value` mobile: `font-size: 16px`, `overflow: hidden`, `text-overflow: ellipsis`

### Loading Overlay
- HTML: `<div id="loadingOverlay">` com 3 dots animados
- Removido sempre via `try/finally` em `initDashboard()` + safety timeout de 10s

---

## Sync Pagar.me — Decisões Técnicas

- **Paginação:** condição de parada é `!paging?.next` apenas (não usar `data.length < PAGE_SIZE` — a API retorna menos que 100 mesmo quando há próxima página)
- **Campo de data:** usar `paid_at` (extraído de `charges[0].paid_at`) com fallback para `created_at`
- **LOOKBACK:** 48h por execução, máximo 10 páginas × 100 registros por run

---

## Campanhas Google Ads

`RISCO_DOC` · `ADDONS` · `RENAVAM` · `ROUBO_FURTO` · `MULTAS`

---

## Custo de Bureau

Fornecedor: Assertiva. Três faixas de preço:
- P1: antes de 2026-02-18
- P2: 2026-02-18 até 2026-03-25
- P3: a partir de 2026-03-26

**Gap residual de ~R$164** entre dashboard e nota Assertiva de março — aceito como estrutural (origem em retentativas de API não rastreadas). Não corrigir.

---

## Decisões Importantes

| Data | Decisão |
|---|---|
| 2026-07-15 | Módulo "Metas" na sidebar: metas migradas das constantes hardcoded p/ tabela `monthly_targets` (migration 013), escrita via Supabase Auth email+senha (signup fechado), leitura anon com fallback em cascata. Meta CAC corrigida p/ R$9 (TARGET_CPA=10 estava errado); upsell 32% |
| 2026-07-13 | P1 completo e em produção: attribution capture no site (gclid/gbraid/wbraid/fbclid/UTMs/fbp/fbc/ga_client_id em `checkouts`, validado com checkout de teste), Click Conversions v24 no dispatch (quando há id de clique; fallback ENHANCEMENT), GA4 client_id real, reconciliação diária (`reconcile-conversion-dispatches`, cron 08:00 UTC). Addon herda atribuição do checkout principal |
| 2026-07-13 | P2.2 no ar: `check-tracking-health` (Telegram, cron horário — failed>10%/24h ou webhook morto). P2.3: PII hasheada em `conversion_dispatches` (migration 011, colunas email/phone dropadas). P2.4: remarketing dinâmico corrigido no GTM (`JS - Contents ID`) e publicado |
| 2026-07-13 | Backlog do reconcile drenado (~148 checkouts pré-webhook); linhas antigas marcadas com `paid_at = created_at`. Pendente: Meta CAPI (P2.1, token), decisão P2.5 (tags primário×upsell), patch stableId begin_checkout (dev), validação gap CAC ≤5% até ~27/07 |
| 2026-07-13 | P0-f: `begin_checkout` alterada de Primária → Secundária no Google Ads (inflava a coluna Conversões; causa principal do overcounting ~25–30%). `purchase` mantida com Contagem "Every" — dedup via transaction ID cobre o disparo duplo web+server |
| 2026-07-13 | P0-e deployado: `computeTransactionId` em `dispatch-conversions` e `retry-conversion-dispatches` (transaction_id de addons = checkout_id, alinhado ao client). Validar na próxima venda de addon |
| 2026-05-01 | Crons migrados do GitHub Actions para cron-job.org |
| 2026-05-01 | `paid_at` adicionado à tabela orders + migration 006 |
| 2026-05-01 | Status `authorized` adicionado aos filtros das views |
| 2026-03-26 | Gap residual Assertiva ~R$164 aceito como estrutural |
| 2026-03-25 | Ticket Médio dinâmico (`receita/pedidos`), removido TICKET hardcoded |
| 2026-03-25 | Colunas Receita e Lucro removidas da tabela de campanhas (eram estimativas) |

---

## Migrações SQL aplicadas

| Arquivo | O que faz |
|---|---|
| `001_canonical_tables.sql` | Cria orders, order_items, google_ads_campaign_daily, bureau_daily |
| `002` – `005` | Índices, views, status authorized |
| `006_paid_at.sql` | Adiciona `paid_at` a orders + backfill + atualiza views |
| `007` – `012` | Tracking gateway (conversion_dispatches, attribution, PII hash) + notificações |
| `013_monthly_targets.sql` | Tabela `monthly_targets` (metas mensais) + RLS anon-read/auth-write + seed jul/26 |

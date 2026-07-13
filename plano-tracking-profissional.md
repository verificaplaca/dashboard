# Plano de Execução — Tracking Profissional VF (P0 → P2)

> Documento de arquitetura para execução por model simples. Cada bloco é autocontido.
> **Contexto fixo:** projeto Supabase do DASHBOARD = `ftmgmfdqdqxboiktxcoj` (este repo).
> Projeto Supabase do SITE = `ozquoloetuzynnyzkado` (repo do dev, "git-hug-you").
> Deploy de function: `supabase functions deploy <nome> --project-ref ftmgmfdqdqxboiktxcoj`.
> SQL do site roda no SQL Editor do projeto do SITE; SQL do dashboard no projeto do DASHBOARD.
> Git: nunca `push --force`; conflito → `git pull --rebase origin main`. Sandbox não commita — devolver comandos git pro usuário.
> Status: P0-c descartado pelo usuário. P0-d (User - DLV → user_data.* no GTM) JÁ FEITO.

---

## P0-a — Retry não deve reprocessar `skipped`

**Problema:** filtro atual trata `meta_status='skipped'` como pendente → toda linha é re-dispatchada até `attempts=5` (confirmado: 69/82 linhas com attempts=5). Risco futuro: quando Meta CAPI ligar, linhas antigas disparariam Purchases retroativos.

**Arquivo:** `supabase/functions/retry-conversion-dispatches/index.ts`

1. Trocar a linha:
   ```ts
   .or('ga4_status.neq.success,ads_status.neq.success,meta_status.neq.success')
   ```
   por:
   ```ts
   .or('ga4_status.in.(failed,pending),ads_status.in.(failed,pending),meta_status.in.(failed,pending)')
   ```
2. No loop, além do guard existente (`if (row.X_status !== 'success')`), NÃO sobrescrever canal cujo status atual é `skipped` com resultado novo `skipped` (evita churn de `updated_at`). Opcional; o guard principal é o filtro acima.
3. Deploy: `supabase functions deploy retry-conversion-dispatches --project-ref ftmgmfdqdqxboiktxcoj`
4. **Validação:** aguardar 30–45min e conferir que `attempts` parou de crescer em linhas com tudo success/skipped:
   `GET /rest/v1/conversion_dispatches_public?select=attempts&order=created_at.desc&limit=20`

---

## P0-b — Timestamps reais (`paid_at`) nos eventos server-side

**Problema:** Meta usa `Date.now()` e GA4 MP não envia timestamp → retries atrasados registram conversão na hora errada. Limites das APIs: GA4 MP aceita eventos retroativos até ~72h (`timestamp_micros`); Meta CAPI até 7 dias (`event_time`). **Verificar esses limites na doc atual antes de confiar.**

**Arquivos:** `supabase/functions/_shared/conversion-dispatch.ts`, `supabase/functions/dispatch-conversions/index.ts`, `supabase/functions/retry-conversion-dispatches/index.ts`, nova migration.

1. Migration `supabase/migrations/009_dispatch_paid_at.sql` (rodar no projeto do DASHBOARD):
   ```sql
   alter table public.conversion_dispatches add column if not exists paid_at timestamptz;
   ```
2. Em `conversion-dispatch.ts`:
   - Adicionar `paid_at?: string | null` à interface `CheckoutConversionData`.
   - `dispatchGA4`: no body, adicionar campo top-level
     `timestamp_micros: data.paid_at ? new Date(data.paid_at).getTime() * 1000 : undefined`.
   - `dispatchMeta`: trocar `event_time: Math.floor(Date.now()/1000)` por
     `event_time: Math.floor((data.paid_at ? new Date(data.paid_at).getTime() : Date.now()) / 1000)`.
3. Em `dispatch-conversions/index.ts`: a RPC já retorna `paid_at` — incluir `paid_at: data.paid_at ?? null` no upsert de `conversion_dispatches` e garantir que `data.paid_at` chega no `dispatchAll` (a interface passa a carregar).
4. Em `retry-conversion-dispatches/index.ts`: passar `paid_at: row.paid_at` no objeto do `dispatchAll`.
5. Deploy das duas functions.
6. **Validação:** próxima venda → conferir linha nova em `conversion_dispatches` com `paid_at` preenchido; no GA4 Realtime/DebugView o purchase server-side deve aparecer com horário coerente.

---

## P0-e — Alinhar `transaction_id` de addon (client × server)

**Problema:** no upsell, o client envia `transaction_id = checkout_id`, mas o server envia `order_nsu` (`addon_...`) → GA4 vê 2 transações (receita duplicada) e o ENHANCEMENT do Ads nunca casa com a conversão client de addon.

**Arquivo:** `supabase/functions/_shared/conversion-dispatch.ts` (só servidor; não mexer no site).

1. Criar helper ao lado de `computeEventId` (mesma regra de prefixo):
   ```ts
   export function computeTransactionId(row: { checkout_id: string; order_nsu: string }): string {
     return row.order_nsu?.startsWith('addon_') ? row.checkout_id : row.order_nsu
   }
   ```
2. `dispatchGA4`: usar `transaction_id: computeTransactionId(data)` (em vez de `data.order_nsu`).
3. `dispatchGoogleAds`: usar `orderId: computeTransactionId(data)`.
4. Deploy: `dispatch-conversions` e `retry-conversion-dispatches` (ambas importam o shared).
5. **Validação:** próxima venda de addon → na tabela `conversion_dispatches` conferir o pedido `addon_...`; no GA4 (Transações) o addon deve aparecer com UMA transação cujo id = checkout_id.

---

## P0-f — Investigar overcounting no Google Ads (~25–30%) — GUIADO, painel do Ads

**Fato observado (29/06–12/07):** conversões reportadas pelo Ads > pedidos pagos Pagar.me todos os dias (ex.: 408 vs 288 em 06/07). Hipótese principal: mais de uma conversion action de purchase marcada como Primária.

Passos no painel (usuário executa, model orienta):

1. Google Ads → **Metas → Conversões → Resumo**.
2. Localizar as duas ações de purchase:
   - `purchase` (tipo Website, id interno `7410940293` — a da tag 97 do GTM e do dispatch server-side)
   - `verificaplaca.com.br (web) purchase` (importada do GA4, id `7410730291`)
3. Conferir a coluna **"Ação de otimização" (Primária/Secundária)**. Se AMBAS estiverem Primárias no mesmo objetivo "Compra" → é dupla contagem. **Fix: deixar só a `purchase` (Website) como Primária; a importada do GA4 vira Secundária.**
4. Conferir se outras ações (begin_checkout, choose_plan, plate_search, print_report, whatsapp) estão como Secundárias. Qualquer uma Primária entra na coluna "Conversões" e infla o número.
5. Na ação `purchase`: conferir **Contagem = "Uma"** (é venda; se estiver "Todas", cada disparo duplicado conta) e a janela de atribuição (30d é razoável; janelas longas + comparação diária explicam parte do gap residual).
6. **Validação:** acompanhar o card "Gap CAC real vs reportado" no Tracking Gateway por 7–14 dias após a mudança; meta = gap ≤ 5%.
7. Registrar no `CONTEXT.md` o que foi mudado (qual ação virou Secundária, data), para o antes/depois do gap fazer sentido.

---

## P1 — Captura de identificadores + Click Conversions + GA4 client_id real + Reconciliação

É a mudança estrutural. Ordem obrigatória: 1 → 2 → 3 → 4 → 5 (o servidor só consegue usar o que o site capturar).

### P1.1 — Patch no SITE (entregar ao dev; gerar arquivo .patch + doc, como nos patches anteriores)

Criar `src/lib/attribution.ts` no repo do site com:

- Na carga do app (import no `App.tsx` ou `main.tsx`):
  - Ler da URL: `gclid`, `gbraid`, `wbraid`, `fbclid`, `utm_source`, `utm_medium`, `utm_campaign`, `utm_term`, `utm_content`.
  - Persistir em `localStorage` (chave `vp_attribution`, com timestamp; sobrescrever se vier clique novo — last click).
  - Guardar também `landing_page` (primeiro `location.href` da sessão) — gravar só se ainda não existir na sessão (`sessionStorage`).
- Função `getAttribution()` que retorna o objeto persistido MAIS leituras no momento da chamada:
  - `fbp`: cookie `_fbp`.
  - `fbc`: cookie `_fbc`; se não existir e houver `fbclid` salvo, montar `fb.1.<timestamp_ms_do_clique>.<fbclid>`.
  - `ga_client_id`: cookie `_ga` — formato `GA1.1.XXXXXXXX.YYYYYYYY` → client_id = `XXXXXXXX.YYYYYYYY` (dois últimos segmentos).
  - `event_source_url`: `location.href`; `client_user_agent`: `navigator.userAgent`.
- No ponto onde o site cria o checkout (chamada à function `criar-checkout-pagarme` — localizar o call site no repo do site), incluir os campos de `getAttribution()` no payload, e a function do site gravar nas colunas novas de `checkouts`.

### P1.2 — Migration no Supabase do SITE (ozquoloetuzynnyzkado, SQL Editor)

```sql
alter table public.checkouts
  add column if not exists gclid text,
  add column if not exists gbraid text,
  add column if not exists wbraid text,
  add column if not exists fbclid text,
  add column if not exists fbp text,
  add column if not exists fbc text,
  add column if not exists ga_client_id text,
  add column if not exists utm_source text,
  add column if not exists utm_medium text,
  add column if not exists utm_campaign text,
  add column if not exists utm_term text,
  add column if not exists utm_content text,
  add column if not exists landing_page text,
  add column if not exists event_source_url text,
  add column if not exists client_user_agent text;
```

### P1.3 — Atualizar a RPC do site

Reeditar `supabase/site_get_checkout_conversion_data.sql` (deste repo) adicionando as colunas novas ao `returns table` e ao `select`, e rodar de novo no SQL Editor do SITE (é `create or replace`).

### P1.4 — Atualizar o dispatch (`_shared/conversion-dispatch.ts`)

1. Estender `CheckoutConversionData` com os campos novos (todos `?: string | null`).
2. **GA4 MP:** `client_id = data.ga_client_id || clientIdSintéticoAtual`. Com client_id real o purchase server-side junta na sessão/atribuição do usuário. (Opcional avançado: capturar também o cookie `_ga_<stream_id>` para `session_id` — deixar para depois.)
3. **Google Ads:** se `data.gclid || data.gbraid || data.wbraid` presente → usar **Click Conversions**:
   `POST https://googleads.googleapis.com/v24/customers/{GADS_CUSTOMER_ID}:uploadClickConversions`
   ```json
   { "conversions": [{
       "gclid": "<gclid>",
       "conversionAction": "customers/<cid>/conversionActions/7410940293",
       "conversionDateTime": "<paid_at formato 'yyyy-MM-dd HH:mm:ss+00:00'>",
       "conversionValue": <valor>, "currencyCode": "BRL",
       "orderId": "<computeTransactionId(...)>"
     }], "partialFailure": true }
   ```
   ⚠️ **Formato não validado ao vivo — conferir nomes de campo na doc atual da API v24 (uploadClickConversions) antes do deploy**, mesmo procedimento do deploy anterior (o endpoint de ENHANCEMENT também foi corrigido assim). Para `gbraid`/`wbraid` o campo substitui `gclid` — confirmar na doc. Fallback: sem nenhum id de clique → manter o caminho atual (ENHANCEMENT).
   Atenção: com Click Conversion registrando a conversão E a tag client-side awct também registrando, o **`orderId` idêntico nos dois é o que deduplica** — por isso o P0-e é pré-requisito. Conferir na doc que a conversion action está com dedup por order_id habilitado.
4. **Meta CAPI:** adicionar ao payload `user_data`: `fbp`, `fbc`; e no evento: `event_source_url`, `client_user_agent` (campo `user_data.client_user_agent`). IP só se um dia for capturado server-side — não bloquear por isso.
5. Persistir os campos novos em `conversion_dispatches` (migration 010 no DASHBOARD com as mesmas colunas relevantes: gclid, gbraid, wbraid, fbp, fbc, ga_client_id, event_source_url, client_user_agent) para o retry não depender de nova chamada de RPC.
6. Deploy das duas functions.

### P1.5 — Reconciliação diária (cobre falha/timeout do Database Webhook)

1. Nova RPC no SITE (`supabase/site_list_paid_checkouts.sql` neste repo; rodar lá):
   ```sql
   create or replace function public.list_paid_checkout_ids(p_since timestamptz)
   returns table (checkout_id uuid)
   language sql security definer set search_path = public
   as $$ select id from public.checkouts where paid_at is not null and paid_at >= p_since; $$;
   grant execute on function public.list_paid_checkout_ids(timestamptz) to anon;
   ```
2. Nova edge function `reconcile-conversion-dispatches` (DASHBOARD):
   - Chama `list_paid_checkout_ids(now() - interval 48h)` no site (headers `BUREAU_SUPABASE_KEY`).
   - Busca os `checkout_id` existentes em `conversion_dispatches` no mesmo período.
   - Para cada faltante: mesmo fluxo do `dispatch-conversions` (RPC `get_checkout_conversion_data` → `dispatchAll` → upsert). Reusar o código extraindo a lógica comum para o `_shared/` se ficar simples; senão duplicar o trecho (função é curta).
   - Limite de segurança: máx. 50 por execução.
3. Job no cron-job.org (mesma conta): POST `.../functions/v1/reconcile-conversion-dispatches`, 1x/dia (sugestão: 08:00 UTC), headers `Authorization: Bearer <SERVICE_ROLE_KEY>` + `Content-Type: application/json`, body `{}`.
4. **Validação:** rodar manualmente 1x via curl e conferir resposta `{ok:true, backfilled:N}`; N deve ser 0 num dia normal.

### P1 — validação geral
Após deploy + site em produção: fazer uma compra teste vinda de um clique com `?gclid=teste` não vale (gclid falso é rejeitado) — validar com tráfego real: conferir em `checkouts` que gclid/fbp/ga_client_id chegam preenchidos, e em `conversion_dispatches` que os campos novos aparecem. No GA4, purchase server-side deve cair na MESMA sessão do usuário (Explorations → transaction_id).

---

## P2 — Itens de consolidação

### P2.1 — Habilitar Meta CAPI
1. Usuário: Events Manager → pixel do VF → Configurações → Conversions API → **Gerar token de acesso**.
2. Secrets no DASHBOARD (Edge Functions → Secrets): `META_PIXEL_ID`, `META_ACCESS_TOKEN`.
3. Nada de deploy — o código já checa os secrets. Recomendado só após P1.4 (fbp/fbc/UA no payload = match quality decente).
4. Teste: usar `test_event_code` do Events Manager (adicionar campo `test_event_code` temporário no body do dispatchMeta, validar no painel Test Events, depois remover).
5. **Dedup a validar:** purchase client (pixel, `event_id` = `evt_...`) + server (CAPI, mesmo `event_id`) devem aparecer como 1 evento no Events Manager. Pré-requisito: build do site em produção com os patches (confirmar com o dev).

### P2.2 — Alerta ativo de saúde do tracking
1. Nova edge function `check-tracking-health` (DASHBOARD), mesmo padrão de notificação da `check-ads-balance` (reusar o mecanismo de e-mail/notificação que ela já usa — ler o código dela antes).
2. Regras: últimas 24h → (a) % de linhas com `failed` > 10%, ou (b) `conversion_dispatches` = 0 enquanto `revenue_daily.paid_orders` > 0 no dia (webhook morto). Qualquer uma → alerta.
3. Cron-job.org: 1x/hora.

### P2.3 — PII: hash at rest em `conversion_dispatches`
1. Migration 011: adicionar `email_sha256 text`, `phone_sha256 text`; migrar dados (`update ... set email_sha256 = encode(digest(lower(trim(email)),'sha256'),'hex')` — requer extensão `pgcrypto`; senão fazer via script) e depois `drop column email, phone`.
2. `dispatch-conversions`: hashear (já existe `sha256Hex` no shared) antes de gravar; gravar só o hash.
3. `conversion-dispatch.ts`: dispatchers passam a aceitar hash pronto (Meta e Google Ads consomem SHA-256 — o phone precisa ser hasheado a partir do formato E.164/só dígitos, manter a normalização atual ANTES do hash).
4. `retry` usa os hashes gravados. A view pública (`008`) não expõe email/phone — conferir que não referencia as colunas dropadas antes de dropar.

### P2.4 — Remarketing dinâmico (GTM): `DLV - Contents` / `DLV - Contents ID`
1. GTM: criar variável JS custom `JS - Contents`:
   `function(){var it={{DLV - items}}||[];return it.map(function(i){return {id:i.item_id,quantity:i.quantity||1}});}`
   e `JS - Contents ID`: `function(){var it={{DLV - items}}||[];return it.map(function(i){return i.item_id});}`
2. Repontar nas tags de remarketing ativas **69, 77, 106** (e órfãs 134–137 por higiene) os parâmetros que hoje usam `{{DLV - Contents}}`/`{{DLV - Contents ID}}`.
3. Validar no Preview (choose_plan, begin_checkout, purchase) e publicar.

### P2.5 — Bug 3 (tags órfãs Primary/Upsell) — exige DECISÃO antes de executar
Contexto: tags 122/123 (awct, labels `-PkUCOaGxPsbELHw-Z1C` primary / `gFgXCKOIxPsbELHw-Z1C` upsell) têm trigger em eventos que o código nunca dispara.
- **Opção A (recomendada):** criar tag awct NOVA "Purchase Split" no trigger `purchase` com `conversionLabel` = Lookup Table sobre `{{DLV - plan_period}}` (`upsell` → label upsell; default → label primary). Deixar as 2 conversion actions como **Secundárias** no Ads (só análise primário × upsell, sem duplicar a coluna Conversões). Deletar/pausar as 12 tags órfãs.
- **Opção B:** abandonar a separação e só deletar as tags órfãs + as 2 actions no Ads.
Perguntar ao usuário antes.

### P2.6 — Cobranças ao dev do site (sem código novo aqui)
1. Confirmar **deploy em produção** do build atual (patches de `items` e `event_id` de upsell já estão no repo, mas build no ar não confirmada).
2. Aplicar o patch pendente do `stableId` em `trackBeginCheckout` (`tracking_event_id_fix.patch`, entregue em 08/07 — ainda não aplicado no repo).
3. Aplicar o patch do P1.1 (attribution) quando pronto.

---

## Ordem sugerida de execução
1. P0-a + P0-b + P0-e (mesma sessão — mesmos arquivos, 1 deploy só)
2. P0-f (painel Ads, guiado)
3. P1.1–P1.4 (patch site → migrations → RPC → dispatch)
4. P1.5 (reconciliação)
5. P2.1 → P2.2 → P2.4 → P2.3 → P2.5 (decisão) → P2.6 (cobranças, pode ser a qualquer momento)

# Instruções — P1.1: Captura de atribuição no site (gclid/gbraid/wbraid/fbclid/UTMs)

Objetivo: capturar identificadores de clique e sessão no client, persistir e
enviar no payload de criação de checkout, pra o dashboard conseguir disparar
**Click Conversions** no Google Ads e **Meta CAPI** com match quality real
(em vez do fallback sintético atual).

## 1. Criar `src/lib/attribution.ts`

Arquivo de referência pronto neste repo: `p1_attribution.ts.example`
(copiar o conteúdo para `src/lib/attribution.ts` no repo do site — o código
não depende de nada específico do projeto, só `window`/`document`/`navigator`
padrão).

Resumo do que o arquivo faz:
- `initAttribution()`: lê `gclid`, `gbraid`, `wbraid`, `fbclid`, `utm_source`,
  `utm_medium`, `utm_campaign`, `utm_term`, `utm_content` da URL atual. Se
  achar qualquer um desses params, **sobrescreve** o objeto salvo em
  `localStorage['vp_attribution']` (regra **last-click**: clique novo
  substitui o anterior, mesmo que o anterior não tenha convertido ainda).
  Também grava `landing_page` em `sessionStorage` — só na primeira vez da
  sessão (não sobrescreve em navegações internas).
- `getAttribution()`: combina o que está em `localStorage` com leituras feitas
  na hora — cookies `_fbp` (Meta Pixel), `_fbc` (Meta Pixel; se não existir e
  houver `fbclid` salvo, monta o formato sintético `fb.1.<timestamp_ms>.<fbclid>`),
  `_ga` (Google Analytics — extrai `ga_client_id` = dois últimos segmentos do
  cookie), mais `event_source_url` (`location.href`) e `client_user_agent`
  (`navigator.userAgent`).

## 2. Chamar `initAttribution()` na carga do app

No `App.tsx` ou `main.tsx` (o mais alto possível na árvore, executado uma
única vez, o mais cedo possível — antes de qualquer lógica que possa limpar
os query params da URL):

```ts
import { initAttribution } from './lib/attribution'

initAttribution()
```

## 3. Incluir os campos no payload de `criar-checkout-pagarme`

**Localizar o call site de `criar-checkout-pagarme`** no repo do site (não
sei o nome exato do arquivo/componente nesta versão do código — buscar por
`criar-checkout-pagarme` ou pela function que dispara a criação do checkout
Pagar.me). No corpo da chamada, adicionar os campos de `getAttribution()`:

```ts
import { getAttribution } from '@/lib/attribution' // ajustar o path conforme o alias do projeto

// ...no ponto em que hoje é montado o payload da chamada a criar-checkout-pagarme:
const attribution = getAttribution()

const payload = {
  // ...campos existentes do payload, não alterar,
  gclid: attribution.gclid ?? null,
  gbraid: attribution.gbraid ?? null,
  wbraid: attribution.wbraid ?? null,
  fbclid: attribution.fbclid ?? null,
  fbp: attribution.fbp ?? null,
  fbc: attribution.fbc ?? null,
  ga_client_id: attribution.ga_client_id ?? null,
  utm_source: attribution.utm_source ?? null,
  utm_medium: attribution.utm_medium ?? null,
  utm_campaign: attribution.utm_campaign ?? null,
  utm_term: attribution.utm_term ?? null,
  utm_content: attribution.utm_content ?? null,
  landing_page: attribution.landing_page ?? null,
  event_source_url: attribution.event_source_url,
  client_user_agent: attribution.client_user_agent,
}
```

## 4. Gravar nas colunas novas de `checkouts`

Na edge function (ou endpoint) que recebe esse payload e faz o `insert` em
`public.checkouts`, incluir os mesmos campos no insert:

```ts
await supabase.from('checkouts').insert({
  // ...campos existentes, não alterar,
  gclid: body.gclid,
  gbraid: body.gbraid,
  wbraid: body.wbraid,
  fbclid: body.fbclid,
  fbp: body.fbp,
  fbc: body.fbc,
  ga_client_id: body.ga_client_id,
  utm_source: body.utm_source,
  utm_medium: body.utm_medium,
  utm_campaign: body.utm_campaign,
  utm_term: body.utm_term,
  utm_content: body.utm_content,
  landing_page: body.landing_page,
  event_source_url: body.event_source_url,
  client_user_agent: body.client_user_agent,
})
```

Pré-requisito: rodar a migration `supabase/site_add_attribution_columns.sql`
(deste repo) no SQL Editor do projeto Supabase do SITE
(`ozquoloetuzynnyzkado`) **antes** de fazer deploy deste código — as colunas
precisam existir antes do insert tentar gravá-las.

## 5. Compatibilidade com fluxo de addon

Se o checkout de addon (`CheckoutAddon`/`Resultado3`) também chama
`criar-checkout-pagarme` ou uma function equivalente, aplicar o mesmo padrão
lá — chamar `getAttribution()` e incluir os mesmos campos no payload.

## 6. Não é bloqueante

Todos os campos são opcionais (`null` se ausente — ex.: usuário chegou por
tráfego direto/orgânico, sem UTM nem clique de ads). O dashboard já trata
esses campos como opcionais no dispatch (fallback pro comportamento atual
quando vierem nulos).

## Checklist de validação após deploy
- [ ] Acessar o site com `?utm_source=teste&utm_medium=teste` e conferir no
      DevTools → Application → Local Storage que `vp_attribution` foi criado.
- [ ] Completar um checkout de teste e conferir na tabela `checkouts`
      (Supabase do site) que as colunas novas vieram preenchidas
      (`utm_source`/`utm_medium` no mínimo; `fbp`/`ga_client_id` dependem de
      cookies do Pixel/GA já terem sido setados antes do checkout).
- [ ] Avisar o time do dashboard quando isso estiver em produção — o P1.4 lá
      passa a usar Click Conversions no Google Ads quando `gclid`/`gbraid`/
      `wbraid` vier preenchido.

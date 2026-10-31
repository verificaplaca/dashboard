# Tracking completo no console · 31/10/2026

O console mostrava apenas uma página de 50 envios com dez campos. A referência
publicada em `https://verificaplaca.github.io/dashboard/dashboard.html` incorpora
`tracking-gateway.html`, que contém os outros indicadores e diagnósticos. Os dois
arquivos remotos foram conferidos e eram idênticos às cópias locais na revisão.

## Informações recuperadas

| Origem | Console |
|---|---|
| Saúde dos últimos 7 dias | Pedidos, envio sem pendências, falhas, qualidade e média de tentativas |
| GA4, Google Ads, Meta CAPI | Sucessos, falhas, pendências e ignorados separados por canal |
| Gráfico 7/14/30/90 dias | Mesmas contagens por dia, tabela acessível e tooltip mobile compartilhado |
| Atribuição dos últimos 7 dias | Google click ID, fbc, fbp e GA client ID |
| Pedidos e filtros | Busca por IDs/valor, status, canal, atribuição e faixa de valor; 50/100/300 por página |
| Detalhe de conversão | Event ID canônico, principal/upsell, match, latência, tentativas, erros e identificadores completos |
| CAC de 30 dias | Custo, pedidos pagos, conversões Ads, CACs ponderados, gap assinado e tabela diária |

`conversion_dispatches_public` é percorrida em páginas de 500, em lotes de até
quatro consultas simultâneas. O corte superior de criação fica fixo durante a
leitura; a ordenação inclui checkout_id para desempate. Nenhum indicador usa só
a página visível. Um erro no meio da leitura rejeita o conjunto inteiro, sem
publicar saúde ou percentuais parciais. Filtros são locais e não alteram os KPIs.
Os registros podem receber atualizações no servidor durante a leitura: trata-se
de consultas REST paginadas, não de uma transação única entre as fontes.

O comparativo financeiro lê `revenue_daily` e `google_ads_campaign_daily`
independentemente do resultado do tracking; uma falha nessa comparação não apaga
o histórico de envios. As campanhas também são paginadas e ordenadas por data e
campaign_id. As leituras usam o loader global e as fontes existentes, somente GET.
Sem migrações, novas permissões, alteração de dispatches ou reenvios.

## Regras preservadas e apresentação

- Janela do histórico: 90 dias. Saúde, canais e atribuição: sete dias corridos.
- A origem aceita success ou skipped como ausência de falha. O console explicita
  que skipped não confirma envio. Contagens de sucesso e ignorados ficam separadas.
- Match forte: quatro ou mais de seis identificadores; médio: dois ou três.
- Event ID esperado: `evt_` + order_nsu. Valores ausentes não viram confirmação.
- A latência preserva os limites de uma e 24 horas usados pela origem.
- O gap conserva `(CAC real - CAC reportado) / CAC real`, inclusive o sinal.
  A legenda antiga invertia a interpretação de subcontagem: o texto foi corrigido.
- Email e telefone mostram somente flags de presença dos hashes. Identificadores
  técnicos vêm da mesma view pública; strings são renderizadas pelo React.
- Mobile usa cards de pedido, filtros e diagnóstico nos modais compartilhados,
  IDs completos com quebra de linha e tabelas financeiras com rolagem local.

## Validação

Testes em `src/domain/tracking.test.ts` confrontam os diagnósticos e os cálculos
com o JavaScript da origem, incluindo skipped, dados ausentes, limites de latência,
conversões fracionárias, filtros e paginação/cancelamento. A conferência da fonte
real leu mais de 22 mil registros únicos, sete dias completos de indicadores e
30 dias de comparação, sem gravar dados. Os números mudam com os novos pedidos.

O arquivo `tracking-gateway.html` anterior permanece intacto. O novo módulo é
React nativo (`LiveTracking.tsx`), sem iframe ou cópia do dashboard inteiro.

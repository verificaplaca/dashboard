# Proteção de reconciliação e investigação de duplicidade — 31/10/2026

**Situação: uma causa de repetição no servidor corrigida; validação integral das três plataformas ainda pendente.** Não concluir que os 38 pedidos repetidos do GA4 foram todos resolvidos. A investigação identificou também ocorrências em navegadores diferentes e falta acesso ao código atual do site para concluir essa parte com segurança.

## Evidência da sobreposição

Em um dos pedidos contados três vezes no GA4, o pagamento aconteceu em 24/10 às 04:00:08 UTC. O registro do Tracking foi criado às 04:00:12,379 e atualizado novamente às 04:00:15,086, mantendo `attempts=1`.

Os logs remotos desse intervalo mostram:

| Rotina | Momento do log de conclusão UTC | Duração |
|---|---|---:|
| Webhook dispatch-conversions | 04:00:12,721 | 2.735 ms |
| Reconciliação | 04:00:15,553 | 5.131 ms |
| Retry | 04:00:16,408 | 665 ms |

A reconciliação iniciou às 04:00:10,549, antes de o webhook terminar. A seleção em lote podia considerar o pedido ausente e enviar a compra novamente. A gravação posterior usava `upsert`, portanto sobrava uma única linha e a coluna de tentativas voltava a 1.

O relatório do mesmo pedido mostra três compras de R$ 14,99: duas com hostname ausente e uma com hostname verificaplaca.com.br; por navegador, duas Safari e uma Chrome. A ausência de hostname é compatível com o payload Measurement Protocol revisado. O cruzamento de horários e a atualização posterior sustentam a correção da sobreposição, sem atribuir todas as demais duplicidades à mesma causa.

## Correção publicada

Somente a rotina de recuperação passou a:

1. Adiar pagamentos com menos de dez minutos, dando tempo para o webhook terminar. Os pedidos continuam elegíveis à janela existente de recuperação de 12 horas.
2. Conferir novamente a linha do pedido após consultar os dados no site. Uma linha válida que tenha aparecido desde a seleção inicial encerra aquela tentativa de recuperação.
3. Na recuperação de metadados quebrados, enviar apenas os canais que ainda não registraram sucesso, preservando o estado de sucesso dos demais.
4. Adiar o pedido quando não for possível confirmar seu estado atual, evitando um reenvio especulativo.

Não se trata de uma garantia geral de exactly-once: não foi implementado um mutex transacional entre todos os emissores nem alterado o fluxo de outro navegador. A proteção cobre a sobreposição confirmada e a seleção em lote desatualizada.

Janela de 12 horas, frequência dos crons, IDs, moeda, valores e payloads de GA4/Ads/Meta foram preservados. Não houve alterações em campanhas, metas, tags, consentimento, configuração de atribuição ou dados históricos. O GTM continua na versão publicada 34, de 28/07. A tag antiga `[API] 6 - Purchase` permanece pausada.

## Confirmações da Meta

Foi acrescentada uma observação da resposta da API após um HTTP de sucesso. Ela registra apenas hash do ID do evento, quantidade recebida, valor/moeda, dia UTC e códigos de diagnóstico. Não registra email, telefone, cookies, URLs de pedido, tokens, mensagens da plataforma ou identificadores originais.

A observação é limitada a 1,5 segundo e não muda o resultado de entrega nem provoca reenvio quando não consegue ler a confirmação.

Na amostra posterior à publicação, foram conciliadas **sete confirmações reais**, totalizando **R$ 152,75**:

- sete respostas com `events_received=1` e HTTP 200;
- sete correspondências com o Tracking e status de sucesso nos três canais;
- zero divergências de valor ou moeda;
- zero hashes de evento repetidos na amostra;
- zero confirmações indisponíveis ou anômalas.

Isso confirma a aceitação dessa amostra pela API da Meta; não certifica seu processamento final, a atribuição ou toda a semana anterior. A diferença histórica de 226 eventos (1.982 na tela da API da Meta versus 2.208 no Tracking entre 24–30/10) permanece sem conciliação individual. A própria interface informa processamento, desduplicação e políticas como fatores que podem afetar suas contagens. Não foram reenviados os 226 eventos por hipótese.

## Validação e implantação

- 11 testes Deno aprovados: sobreposição de pagamento recente, leitura desatualizada, recuperação legítima, erro de leitura, preservação de canais concluídos, IDs/valores e confirmação da Meta.
- `deno check` aprovado na rotina de reconciliação e diff sem erros de formatação.
- Código das funções baixado após publicação e comparado byte a byte com o código testado.
- `dispatch-conversions`: versão 15 → **16**, ACTIVE, JWT preservado.
- `reconcile-conversion-dispatches`: versão 6 → **7**, ACTIVE, JWT preservado.
- `retry-conversion-dispatches`: continua na versão **13**.
- Painel publicado atualizado: 2.270 pedidos na janela móvel de sete dias, zero falhas/pendências, os três canais com sucesso. Essa janela difere da semana fechada usada na auditoria.
- Google Ads consultado novamente: o último diagnóstico disponível da API para a ação purchase continua `EXCELLENT`, 302/302, sem pendências; é um resumo anterior à implantação, portanto não foi apresentado como confirmação imediata de todos os novos pedidos.

## Pendência para concluir o GA4

O site publicado possui proteção local por navegador (`vp_purchases_fired`), mas essa proteção não é compartilhada entre Safari, Chrome e outros dispositivos. A compra do GA4 usa o ID correto, porém um ID correto não resolve sozinho todos os casos envolvendo identidades distintas.

É necessário revisar o código atual do site e o broadcast de pagamento, preservando a identidade canônica da compra e os dados de navegação/atribuição. O Supabase do site (`ozquoloetuzynnyzkado`) retornou HTTP 403 ao tentar ler a função processar-callback com a conta disponível. Foi solicitado ao usuário o repositório atual ou o contato de quem mantém esse código. Não se aplicou um snapshot antigo nem se desligou o purchase do navegador às cegas.

A validação final deve usar compras novas depois dessa correção, comparar por pedido os valores e a contagem no GA4 e acompanhar processamento da Meta e diagnóstico do Ads após a atualização de seus relatórios. Os 39 registros extras e R$ 648,44 identificados na auditoria histórica não foram apagados.

## Reversão

As versões anteriores foram baixadas antes da publicação, em `node_modules/.cache/tracking-audit-20261031/production` no checkout canônico. Para reverter, republicar somente dispatch-conversions e reconcile-conversion-dispatches a partir desse diretório, mantendo o projeto ftmgmfdqdqxboiktxcoj e JWT habilitado. O código anterior também está no commit-base 590622d. Não há migração de banco ou mudança em secrets para desfazer.

Evidência visual do painel após atualização: `docs/evidencias/tracking-protecao-2026-10-31.jpg`.

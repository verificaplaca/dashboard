# Simulador manual — 02/11/2026

A tela inicial do Simulador funciona como uma calculadora de cenários. A base de 7, 14 ou 30 dias completos preenche verba média, CAC, ticket e custo de bureau. A carga inicial e Restaurar cenário reproduzem a referência histórica, sem partir de uma etapa recomendada ou de um orçamento fixo.

O usuário altera investimento diário, CAC e horizonte. O CAC é mantido quando a verba muda; receita, pedidos, bureau, impostos, taxas, margens e lucro acompanham as premissas. A edição do CAC pelo card também atualiza o campo superior. Editar ticket/bureau não altera a eficiência. Edições de transações/conversão continuam definindo pedidos por real investido, como explicitado no editor. Restaurar cenário limpa essas hipóteses e volta à base selecionada.

## Conta

- Pedidos no período = verba diária × dias ÷ CAC escolhido.
- Receita = pedidos × ticket; bureau = pedidos × custo por pedido.
- Imposto usa as metas vigentes por data. Taxas cadastradas na DRE acompanham a receita; as demais despesas conhecidas são mantidas no horizonte, conforme o comportamento anterior.
- A comparação usa a verba média e a eficiência observadas na base, sem aplicar as alterações hipotéticas ao controle histórico.
- O CAC segue a definição existente do painel (Ads ÷ pedidos pagos), não clientes novos únicos. Quantidades simuladas podem ser fracionárias; valores monetários são arredondados por dia.

A calculadora manual não aplica elasticidade, perda de adaptação ou sazonalidade automática. A curva tem somente a eficiência escolhida; removidos os cenários automáticos, a estrela/sugestão de orçamento de maior lucro e o indicador automático de aumentar/manter/reduzir da tela inicial. Os modelos antigos e cenários salvos continuam no acesso explicitamente denominado “Modelos opcionais e cenários salvos”, sem influenciar a calculadora manual.

Sem todas as despesas da DRE, o resultado continua identificado como parcial. Não são inventados custos ausentes. Não há alteração de fontes reais, permissões, dados do site, campanhas ou orçamento do Google Ads.

## Validação

218 testes de domínio/componentes e 12 de paridade aprovados (230 no total). Seis novos testes verificam proporcionalidade com CAC constante, CAC maior, conservação do CAC ao mudar verba/ticket/bureau, despesas fixas versus variáveis, investimento zero, curva única e rejeição de eficiência inválida.

Inspeção local do componente real com fontes agregadas de produção, no acesso demonstrativo do shell (sem autenticação ou gravações reais):

- Base 19/10–01/11, R$3.105,04/dia, CAC aproximadamente R$9,0886: o cenário inicial coincide com manter a referência.
- Verba R$6.200/dia conserva o CAC e aumenta as transações/receita/bureau proporcionalmente.
- CAC R$12 com a mesma verba reduz pedidos para 15.500 no horizonte de 30 dias.
- Editar CAC pelo card para R$10 e informar R$6.000/dia produz 18.000 pedidos e R$180.000 de mídia, sincronizando o campo superior.
- Restaurar cenário volta à base e à indicação “Mesmo resultado que manter”.
- Em viewport de 390×844, os dois inputs têm as mesmas bordas (16–368 px), sem overflow da página. Viewport de teste restaurado.

Evidências: `docs/evidencias/simulador-manual-desktop.jpg` e `simulador-manual-mobile.jpg` (prévia local; hipóteses digitadas para teste, não resultados realizados).

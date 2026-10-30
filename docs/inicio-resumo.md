# Início · resumo da operação

A inicial de produção (`#/inicio`) passa a ter composição própria em `OperationsHome.tsx` e CSS isolado em `operations-home.css`. Não monta o painel legado, seus controles, gráficos e cards dentro da página. O componente antigo `ExecutiveSummary` permanece preservado para os usos existentes; a Visão Geral e o motor original não são alterados.

## Leitura da página

- Cabeçalho compacto com saudação, data de São Paulo e atualização das fontes da inicial.
- Quatro indicadores: receita, resultado gerencial/resultado parcial, pedidos pagos e CAC por pedido pago.
- Gráfico de barras dos até sete dias mais recentes do período financeiro, com valores e alternativa textual acessível.
- Meta de receita mensal com anel de progresso, valor-alvo e acesso às metas. Percentual pode exceder 100%; o anel visual é limitado a uma volta.
- Prioridades reais, pendências, tarefas atrasadas e vencimentos do operador.
- Investimento em Google Ads, orçamento mensal e resumo de tracking.
- Acessos rápidos aos módulos operacionais; Usuários aparece exclusivamente para administrador. Início também passa a ter entrada própria no menu lateral, distinta de Visão Geral.

## Fontes e limites

O resumo financeiro usa `useFinanceSource`, `useExpenseMonth` e `buildDre`, os mesmos leitores e cálculos do Financeiro. Mostra o mês do último dia encerrado, com intervalo explícito; hoje não é tratado como um dia completo. No primeiro dia do mês, mostra o mês anterior fechado. Portanto, comparar com a Visão Geral exige selecionar o mesmo intervalo. A inicial não projeta receita ou lucro futuro.

Resultado inclui as despesas privadas disponíveis. Categorias ausentes, imposto em fallback ou despesas sem acesso impedem o rótulo de resultado completo. Campos não disponíveis não aparecem como zero. Uma falha financeira não impede consultar tarefas, metas disponíveis e atalhos. Receita e CAC não recebem fórmulas novas: CAC é mídia realizada dividida pelos pedidos pagos; sem pedidos não existe divisão por zero.

Metas de receita, CAC e orçamento vêm da tabela `monthly_targets`, com herança individual do último campo não nulo até o mês selecionado. Nunca herda mês futuro nem substitui zero explícito. Ausência/falha é indicada; não inventa meta. A meta de receita é do mês inteiro, mesmo com o realizado até ontem.

Tracking consulta somente os 50 envios mais recentes nos sete dias de calendário indicados. Contagens são por evento: falha tem prioridade sobre pendência; status desconhecido fica a conferir. `skipped` não é tratado como confirmação de entrega. A amostra não é apresentada como total da operação ou taxa global de sucesso. Tarefas continuam privadas e obedecem à conta autenticada. A atualização geral também atualiza o resumo das tarefas; a consulta silenciosa existente permanece.

O loader global, seu fade e os temas existentes continuam em uso. Não há novas escritas no banco, alteração de contas ou ação em campanhas nesta revisão.

## Validação

Testes específicos cobrem herança por campo, mês futuro, zero explícito, metas inválidas e contagem de tracking sem duplicar eventos ou confundir skipped/unknown. Os testes existentes de DRE e paridade preservam os contratos financeiros.

Validação concluída: 166 testes de domínio e 12 de paridade aprovados (178 no total), build TypeScript/Vite aprovado. Revisão visual da versão compilada em desktop e em 390×844, nos temas claro/escuro, com atualização pelo loader global e tratamento de indisponibilidade financeira. Os fixtures locais usam conta e tarefas fictícias, sem criar sessões reais ou executar escritas no backend.

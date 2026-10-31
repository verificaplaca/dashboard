# Painel publicado e tratamento de indisponibilidade — 31/10/2026

## Verificação com dados reais

Conferência autenticada em `https://verifica-placa-ops.vercel.app/`, com viewport mobile. A aba aberta inicialmente mantinha um bundle de uma versão anterior; a verificação continuou após recarregar a versão publicada.

Conferidos Início (indicadores, receita, meta, tarefas, canais e atalhos), Visão Geral (resultados, gráficos, eficiência, produtos, funil, operação e estornos), Financeiro (resumo, DRE, despesas e fontes), Campanhas (tabela, seleção de campanha, Search, distribuição e gráficos), Tracking (registros, canais e filtro), Metas (campos e histórico), Usuários (lista e abertura do editor), Tarefas (lista, detalhe, responsável, prazo e histórico) e Perfil.

No simulador: cartões e curva com base real, Decisão de hoje e critérios expandidos, Plano até janeiro e comparações, Investimento ideal com subseções expandidas e Cenário livre com custos, ciclos, gráficos e tabelas. As projeções continuam identificadas como hipóteses.

As consultas financeiras identificaram a origem Supabase e atualização do próprio dia. A verificação não salvou usuários, senhas, metas, despesas, tarefas, comentários, campanhas ou hipóteses do usuário. Filtros de campanha/tracking foram restaurados após a consulta.

O widget de tracking no Início apresentou uma falha pontual de leitura. A consulta REST equivalente respondeu HTTP 200 com os 50 registros solicitados na repetição; a página de Tracking também abriu com os envios reais.

## Faixa vermelha

A faixa vinha de `demoDataBanner`, criado pelo motor legado quando uma leitura falhava e `makeDays()` produzia números sintéticos. Em uma atualização que falhava após uma leitura válida, o motor também podia manter `usingRealData=true`; portanto, confiar apenas nessa flag não era suficiente.

O adaptador publicado agora bloqueia a publicação desse estado antes de liberar os indicadores, descarta o controlador e seus gráficos e limpa o resumo. A faixa fica suprimida e a regra que a tornava vermelha e visível foi removida. Em uma indisponibilidade, aparece somente uma área neutra com nova tentativa; os indicadores não são preenchidos por exemplos. A nova tentativa recria a leitura para não carregar o marcador da falha anterior.

A fonte congelada é preservada para os testes de paridade financeira. A proteção é aplicada na integração usada pelo painel, sem alterar as fórmulas ou as consultas analíticas.

As leituras públicas de metas e tracking ganharam prazo por tentativa e uma repetição automática para falhas transitórias de rede, HTTP 429 ou 5xx. Erros de autorização e respostas malformadas não são repetidos; cancelamentos continuam respeitados.

## Validação da proteção

Testes unitários cobrem dados reais, fallback em atualização com flag antiga de origem real, estado sintético sem marcador, leitura pendente, recuperação, repetição transitória, autorização, cancelamento e limite de tentativas.

Em ambiente local isolado, foram provocadas falhas da fonte na abertura e após uma leitura válida. Em ambos os casos: nenhum banner, nenhum KPI sintético, loader encerrado e ação de nova tentativa disponível. No resumo, os quatro indicadores ficaram indisponíveis em vez de receber exemplos. A leitura voltou após retirar a falha e tentar novamente.

O teste de falha controlada ficou restrito ao ambiente local. O painel publicado foi consultado com suas fontes reais.

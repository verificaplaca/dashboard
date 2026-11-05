# Revisão de integração e rolagem · 04/11/2026

Revisados os módulos publicados: Início, Visão Geral, Financeiro, Simulador, Campanhas, Tracking, Metas, Usuários, Tarefas e Perfil. Nenhum iframe no DOM desses módulos. Visão Geral e Campanhas usam Shadow DOM no mesmo documento; o gerador remove os iframes do modelo original e preserva os cálculos.

Encontrada rolagem vertical interna de 18px em `.financial-scroll` na Visão Geral: `overflow-x:auto` fazia o eixo vertical computar como auto. As grades da Visão Geral e do Simulador também exigiam largura mínima de 1100px. Removidos os contêineres de rolagem dos cards e a largura mínima; grades com três colunas em larguras intermediárias e layout de celular preservado. Removida a instrução de rolar lateralmente nos cards do simulador.

Menu lateral, tabelas largas, navegação por abas, diálogos, seletores e detalhes temporários mantêm sua rolagem funcional. Não são páginas embutidas. A caixa de atendimento com mensagens e os demais módulos em demonstração não estão habilitados na operação publicada.

O dashboard legado independente e a cópia de referência ainda contêm embeds históricos de Google Ads CAC e Tracking. Eles não são usados nas rotas do novo painel; o modelo nativo exclui essas tags. Não se alterou a referência congelada nem o motor de cálculos.

Validação: testes de domínio e paridade, build de release e build Vercel; inspeção do DOM com Shadow DOM e dos contêineres com overflow real. Evidências em `output/revisao-rolagem-2026-11-04/`.

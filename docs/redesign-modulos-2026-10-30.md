# Redesign dos módulos — 30/10/2026

Extensão da linguagem visual aprovada para Início, Financeiro, Simulador (simples e avançado), Campanhas, Tracking, Metas, Tarefas e Perfil. Superfícies claras, espaçamento consistente, controles arredondados, tabelas legíveis, estados sem dados e formulários padronizados. Demonstração recebe os componentes compartilhados.

Campanhas incorpora folha de estilo própria no Shadow DOM, descrição de estimativas e ordenação por teclado. Valores dos gráficos são acessíveis pela mesma tabela existente. Tracking diferencia estados desconhecidos/ignorados de falhas. Visão Geral e seus 15 cards permanecem preservados.

Validação: npm test; npm run release; npm run build:vercel; git diff --check. Inspeção visual local de Financeiro, Simulador, Campanhas, Metas (desktop e 390px), Tracking e estado não autenticado de Tarefas. Não foram efetuadas gravações autenticadas nem alterações no banco. Motor de cálculo original preservado por testes de paridade.

Publicação: push main pelo remote SSH github-verificaplaca aciona Vercel e GitHub Pages. Artefatos console/ gerados pelo script release.

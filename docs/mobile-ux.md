# UX mobile do painel — 30/10/2026

## Escopo

Revisão dos módulos ativos: Início, Visão Geral, Financeiro, Simulador simples e suas quatro áreas avançadas, Campanhas, Tracking, Metas mensais, Usuários, Tarefas e Meu perfil. Inclui entrada no painel e recuperação de senha.

## Comportamento

- Até 780 px, cabeçalho de uma linha, identificação do módulo e menu lateral com nomes. O menu bloqueia a rolagem e o conteúdo ao fundo, mantém o foco, fecha com Escape e devolve o foco ao botão de abertura.
- Controles principais com área de toque de 44 px; campos com fonte de 16 px para evitar zoom automático no iOS. Respeito às áreas seguras e à altura dinâmica da tela.
- Até 600 px, usuários e eventos de tracking aparecem em cartões; histórico de metas usa detalhes expansíveis. As mesmas ações e permissões são mantidas.
- Tarefas abre em lista no celular, com cartões e filtros recolhíveis. O quadro continua disponível com rolagem por coluna; a lista não inicia arraste de cartões.
- Indicadores da Visão Geral ficam em duas colunas no celular e três no tablet. Dicas dos indicadores abrem em diálogo no mobile, sem sobrepor números. Gráficos cabem na largura disponível; tabelas analíticas conservam a comparação entre colunas, com rolagem local, primeira coluna fixa e indicação de gesto.
- Financeiro mantém valores monetários inteiros em uma coluna no celular. Simulador reorganiza controles e cartões, preservando as premissas e as contas.
- Diálogos de período, despesas, usuários e tarefas respeitam a altura da tela e permitem rolar o conteúdo.

## Implementação

`console-app/src/mobile-ux.css` contém os ajustes da aplicação React. A Visão Geral usa Shadow DOM e recebe `console-app/public/overview/mobile.css` pelo gerador `scripts/build-overview.mjs`. A fonte congelada do dashboard e suas fórmulas não foram editadas. A recuperação de senha tem ajustes próprios em `password-recovery.css`.

## Validação

Revisão visual e interações em Chrome com viewport de 390×844, conferências em 320×568 e 768×1024 e regressão desktop. Exercitados menu/fechamento/foco, seletor de período, abas avançadas, filtros de tarefas e abertura de formulários de despesas/usuários/tarefas, perfil, login e formulário de recuperação. Gerente conferido com metas desabilitadas.

A revisão local utiliza dados fictícios com nomes e identificadores longos, diferentes estados de envios e tarefas, sem gravar registros, redefinir senhas ou enviar emails. A simulação de dados permanece fora do repositório e da versão publicada. Testes de domínio e paridade do motor financeiro, build TypeScript/Vite e verificação do artefato de publicação complementam a revisão visual.

A emulação valida disposição e navegação; não substitui testes físicos com teclado virtual e Safari em iPhone.

## Revisão complementar de 31/10/2026

Auditoria ampliada, incluindo subseções e detalhes expandidos, casos de conteúdo longo e falhas. Matriz de cobertura e correções: [auditoria por página e seção](mobile-ux-auditoria-2026-10-31.md).

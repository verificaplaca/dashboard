# Auditoria mobile por página e seção — 31/10/2026

Segunda revisão do painel, percorrendo também as seções inferiores e os detalhes recolhidos. Escopo: todas as rotas ativas da operação, perfil, acesso e as nove páginas em preparação acessíveis pela busca. Os fluxos fictícios da demonstração separada não fazem parte da operação real.

## Cobertura

| Página | Seções e estados conferidos |
|---|---|
| Início | Indicadores, receita diária, objetivo mensal, tarefas e prazos, Google Ads, tracking, aviso da DRE, atalhos, período personalizado e aplicação do filtro compartilhado. |
| Visão Geral | 15 indicadores e suas informações, resultado bruto/líquido diário, comparação de períodos, investigação por dia e valores dos gráficos, receita/custos, sete dias da semana, quatro medidores de eficiência, meta de resultado, dois produtos e respectivas referências, funil, conversões por campanha, saldo e autonomia, histórico de saldo, bureau e fornecedores, três indicadores de estornos. |
| Campanhas | Seleção de campanha, datas e período personalizado, ordenação e rolagem da tabela, oportunidades Search, diagnóstico por campanha, causas de perda de visibilidade, distribuição do investimento, evolução das impressões perdidas e tabela de valores do gráfico. |
| Financeiro | Mês, resumo, todas as linhas da DRE, resultado, despesas complementares, estornos, fontes e formulário de despesas com rolagem e botões acessíveis. |
| Simulador simples | Base e horizonte, investimento, adaptação, 15 cartões, edição de ticket e recálculo, resultado, curva, três hipóteses e explicações expandidas. |
| Decisão de hoje | Indicador, critérios e acumulado, etapas, calendário, referência de investimento, histórico e formulários de registro/maturação expandidos. |
| Plano até janeiro | Avaliação real, hipótese de aprendizado, quatro bases, condições de comparação, valores exatos e recuperação recente. Conferido com dados disponíveis e com base indisponível. |
| Investimento ideal | Marco de alterações, controle, faixa da busca, cinco bases, resultado, curva e valores, decomposição do lucro, sensibilidade da resposta, antes/depois e comparação de CAC. Todos os detalhes expandidos. |
| Cenário livre | Base, premissas e custos, lucro consolidado, ciclos, comparação financeira, gráficos acumulado/por ciclo e investimento/lucro com tabelas de valores, sensibilidade, resultado por ciclo. Salvar e reabrir uma hipótese local com nome longo. |
| Tracking | Filtros, estados dos canais, identificadores longos, tentativas, rodapé e paginação; lista vazia. |
| Metas | Todos os campos, revisão anterior à gravação, histórico e abertura de mês; gerente com edição desabilitada. |
| Usuários | Guia de perfis, busca, cartões com nomes/emails longos, criação, edição e revisão, senha manual e envio de link (somente abertura dos diálogos), situação suspensa e falha de carregamento. Gerente bloqueado na rota. |
| Tarefas | Indicadores, escopos, filtros, lista, quadro com rolagem local, detalhe, responsável, prazo, estado, edição, nova tarefa, comentários e histórico com textos/links longos. |
| Perfil e acesso | Identidade, campos e ações, entrada no painel, acesso restrito, link de recuperação inválido e formulário de nova senha válido. |
| Páginas em preparação | CRM, Clientes, Vendas, Consultas, Caixa compartilhada, WhatsApp, Evolução da equipe, Suporte interno e Administração: título, aviso de integração e ações. Conferidas em 320 px. |
| Navegação e diálogos | Menu, busca com e sem resultados, atalho de busca com o menu aberto, identificação da página, Escape, foco e retorno do foco, cabeçalho fixo, conteúdo longo, botões e altura disponível. |

## Problemas encontrados e corrigidos

- A orientação horizontal do celular acionava a navegação de desktop. O modo compacto agora também considera telas baixas em orientação horizontal.
- O calendário de Campanhas fechava no próprio clique de abertura. A integração agora impede a propagação desse clique, posiciona o painel fora dos contêineres que o recortavam, mantém os campos personalizados visíveis e fecha com Escape.
- O calendário podia ficar atrás de controles do cabeçalho; sua camada e o posicionamento no tablet foram corrigidos.
- Valores de produtos ficavam cortados em 320 px. Os cartões passam a uma coluna nessa largura.
- A análise semanal exigia rolagem lateral e tinha valores apertados. Agora usa grade no celular e uma coluna em telas menores.
- As informações dos medidores de eficiência podiam sobrepor cartões. Usam o mesmo diálogo de ajuda dos demais indicadores no celular, com área maior para toque.
- Descrições e números da DRE podiam encostar em 320 px. As linhas passam a exibir descrição e valor empilhados, preservando os papéis acessíveis da tabela.
- O resultado do investimento ideal e o consolidado do cenário livre tinham valores sobrepostos. Os blocos e indicadores da base foram reorganizados.
- Emails longos extrapolavam os diálogos de usuários. Os textos agora quebram dentro do conteúdo, e a revisão de alterações apresenta nome, email, perfil e situação completos.
- O tema escuro tinha regras antigas que não alcançavam o conteúdo em Shadow DOM. A conversão de seletores agora cobre os formatos existentes, com testes de regressão, e os avisos ganharam contraste adequado.
- Avisos de confirmação ficavam estreitos e altos. Agora usam a largura disponível, com botão de fechar acessível.
- A busca agora informa quando não encontra um módulo. O atalho de teclado fecha o menu antes de abrir a busca, evitando duas retenções de foco simultâneas.
- Tabelas extensas têm indicação de rolagem e regiões que recebem foco, inclusive em subseções montadas depois de expandir detalhes.

## Método e limites

Chrome com emulação de 320×568 e 390×844, conferência em 768×1024, orientação horizontal de 844×390 e formulário com 390×480 de altura disponível. A rodada final percorreu as dez rotas em cada uma das quatro configurações (40 verificações), além das quatro áreas avançadas do simulador com todos os detalhes abertos. Conferência de largura do documento, conteúdo dentro de cada seção, campos, tabelas, rolagem e abertura/fechamento dos estados interativos. As tabelas que comparam várias colunas mantêm rolagem interna intencional.

Dados fictícios servidos somente no ambiente local, incluindo nomes, emails, identificadores, comentários longos, listas vazias e erros. Nenhuma senha foi digitada, nenhum email enviado, nenhuma meta ou despesa real salva, nenhum usuário criado e nenhuma campanha alterada. A hipótese salva para teste pertence apenas ao navegador local de revisão.

Build TypeScript/Vite, 179 testes de domínio e os 12 testes de paridade confirmam a preservação do motor financeiro original. A emulação não substitui um teste em iPhone físico com Safari e teclado virtual.

# Console — primeira integração de produção

Fonte canônica: `console-app/`, no repositório `verificaplaca/dashboard`.
Painel principal: https://verifica-placa-ops.vercel.app/ (Vercel, deploy automático da main).
Saída publicada: `console/`, pelo GitHub Pages existente (main, raiz).
O endereço original e `dashboard.html` permanecem disponíveis.

## Escopo entregue

- Inicial com saudação pelo horário de São Paulo, resumo financeiro real e acesso aos módulos.
- Visão Geral com o motor original integral, sem iframe. Mantém os 15 indicadores, filtros, impostos, herança de metas e comparativos.
- Simulador em `#/simulador`, disponível na sidebar real e na busca: decisão com CAC/lucro realizados, comparação gradual/direto até janeiro, otimizador e cenário livre. Fontes agregadas somente leitura, cálculos condicionais e hipóteses editáveis.
- Campanhas com o painel original nativo, incluindo perdas de participação de impressão.
- Tracking em leitura da view pública existente, paginado em 50 registros e com janela de 90 dias. Não dispara reenvios ou reconciliação.
- Login Supabase com a conta existente; identidade verificada pelo servidor, sessão em sessionStorage e renovação pelo SDK. Sem seletor de papéis em produção.
- Metas organizacionais mensais e imposto em `monthly_targets`, com confirmação de valores e controle de concorrência por `updated_at`. Meses novos usam INSERT; não substituem silenciosamente registros concorrentes.
- Perfil com nome de exibição salvo no Supabase Auth. Metadata não concede permissões.
- Tarefas reais em `console_tasks`: quadro/lista, busca e filtros no servidor, criação/edição, responsáveis, prioridades, prazos pelo calendário de São Paulo, comentários, histórico, conclusão/reabertura e arquivamento/restauração. Resumo de pendências na inicial.
- Demonstração completa preservada em `?demo=1`, com armazenamento e comandos fictícios separados do modo de produção.

A primeira integração preservou o esquema original. Em 08/10/2026, as migrações aditivas 017 e 018 criaram exclusivamente o backend de tarefas e suas permissões. Não alteram financeiro, pedidos, cron ou sync. Login, edição de perfil e de metas só ocorrem por ação do usuário autenticado. O banco mantém a regra existente: qualquer usuário `authenticated` autorizado pela RLS pode editar metas. Isso ainda não representa os papéis e equipes da demonstração.

## Contratos que faltam para a operação completa

CRM, responsabilidade por clientes, suporte, chat compartilhado, instâncias WhatsApp, convites, RBAC por equipe e progressão individual não têm tabelas/APIs próprias no backend auditado. Permanecem em integração no modo real; o botão de demonstração permite revisar a experiência já aprovada.

Pedidos sincronizados usam `(provider, provider_order_id)`; tracking usa `checkout_id`. Não inventar vínculos entre as duas identidades. Não usar `raw_json`, emails e telefones nas leituras públicas do console. O projeto do site (`ozquoloetuzynnyzkado`) tem checkouts e consultas; o projeto analytics (`ftmgmfdqdqxboiktxcoj`) reúne financeiro e tracking. As credenciais privilegiadas do site devem permanecer no servidor.

Próxima etapa: definir provedores WhatsApp, contratos de consulta/entrega e backend de operação com autorização, persistência, auditoria e eventos reais para gamificação. Não converter dados fictícios em registros de produção.

## Preservação e verificação

Em 07/10/2026, 18:40:35 UTC, auditoria somente leitura do projeto analytics confirmou 48.306 pedidos, 48.306 itens, 18.909 dispatches e 3 meses de metas. Contagens são uma fotografia de uma base que continua recebendo syncs, não um backup restaurável.

O SHA-256 de `dashboard.html` e da referência congelada é `9d2852deed6eb03200c4836adf86bc3de1e31ab61543497a0834553674bfdedb`. O release é bloqueado se a fonte canônica divergir da referência: revisar e recapturar deliberadamente antes de atualizar o motor.

49 testes cobrem o domínio demonstrativo, a apresentação, equivalência de fontes/cálculos, validação dos valores mensais, rejeição de identidade ausente, concorrência e falha de RLS. Testes de escrita no frontend usam cliente controlado. `supabase/tests/console_tasks.sql` verifica as RPCs e a RLS no PostgreSQL com fixtures exclusivamente transacionais e ROLLBACK, incluindo acesso anônimo, membros externos/inativos, identidade trocada, concorrência e preservação do histórico. Não deixa usuários ou tarefas fictícias persistidos. O login e a persistência de edição precisam de validação com a conta do operador, sem pedir senha em chat.

## Desenvolvimento e publicação

Node 22, React 19, TypeScript, Vite e Chart.js nas versões da referência.

```sh
npm --prefix console-app ci
npm --prefix console-app run dev
npm --prefix console-app test
npm --prefix console-app run release
```

`release` prepara exclusivamente a pasta gerada `console/`. Commitar fonte e saída juntas. CI em push na main, PR e execução manual reconstrói, testa e compara os arquivos gerados; não recria schedules de sincronização.

O frontend usa caminhos relativos para funcionar em `/dashboard/console/` e outros subdiretórios. Demo, páginas, gráficos e SDK Auth ficam em chunks separados. O SDK Auth só é solicitado ao entrar ou restaurar uma sessão; a demonstração não é executada no modo real.

Rollback do frontend: reverter o commit que adiciona/atualiza `console/`. O dashboard original continua no endereço habitual. Não reverter metas ou perfil editados posteriormente pelo operador durante um rollback de frontend; tais alterações pertencem ao banco existente.

## Tarefas — autorização e manutenção

Tarefas exigem login; não existe leitura pública. `console_task_members` é uma lista explícita de contas habilitadas, separada dos papéis simulados. O único operador já existente foi habilitado como gestor; nenhuma nova conta Auth foi criada. Gestor vê todas as tarefas. Membros veem somente as que criaram ou receberam. Responsável pode mudar o status e comentar; criador/gestor edita detalhes e arquiva/restaura. Metadata do perfil não concede essas permissões.

O browser tem apenas SELECT sob RLS e EXECUTE nas RPCs públicas. Escritas diretas, autoelevação de papel e DELETE são bloqueados. Implementações ficam no schema privado. Cada RPC verifica `p_actor` contra `auth.uid()`, mantém a membership bloqueada contra revogação concorrente, valida o responsável ativo e grava a auditoria na mesma transação. Atualizações exigem a versão observada; criação usa UUID estável para evitar duplicação em repetição de envio. O editor permite consultar a versão atual e manter o rascunho antes de tentar novamente.

Comentários são texto simples. O histórico registra ator, campos antes/depois, horário de execução e versão; só participantes autorizados podem lê-lo. Arquivar é reversível e não apaga eventos. Listas/histórico carregam 50 registros por vez, com contagem exata e opção de carregar mais; os indicadores vêm de agregação protegida por RLS. Atualizações são consultadas a cada 30 segundos enquanto a página está visível e ao voltar à janela.

Para contas futuras, cadastrar membership deliberadamente pelo backend administrativo (`service_role`), nunca pelo browser ou por metadata. Antes de criar novas contas Auth para tarefas, revisar a política financeira **já existente** de `monthly_targets`, que permite escrita a qualquer `authenticated`. Esta entrega não provisiona novos usuários nem modifica essa política. Convites/equipes serão integrados no módulo de administração.

Validado: 49 testes automatizados; testes SQL transacionais de autorização e fluxo; validação visual isolada com fixtures de quadro/lista, criar, comentar, mudar status, conflito com rascunho preservado, arquivar/restaurar e popup a 390px, sem overflow horizontal da página. A validação visual autenticada utiliza um harness temporário separado, que não é incluído no release. Login e ações pela conta real do operador continuam disponíveis para validação humana, sem geração de sessões de impersonação.

Rollback: reverter frontend e manter as tabelas/permissões privadas preserva tarefas criadas pelo operador. Nunca apagar `console_tasks` ou o histórico para reverter uma versão da interface. Migrações 017/018 foram aplicadas individualmente pelo Management API do CLI no projeto principal; não executar `db push` em bloco sobre o histórico legado.


## Simulador integrado à publicação principal — 26/10/2026

O módulo estava disponível apenas no console separado hospedado no ChatGPT. Foi incorporado à entrada real `Production.tsx`, à navegação principal (abaixo de Visão Geral), à busca e à rota `/dashboard/console/#/simulador` do GitHub Pages. A demonstração também oferece o acesso com sua permissão de Visão Geral. O estado visual usa `UIProvider`, sem inicializar dados fictícios no modo real.

Inclui as quatro áreas já validadas: Decisão de hoje, Plano até janeiro, Investimento ideal e Cenário livre, com gráficos de lucro líquido, comparação gradual/direto/controle, CAC constante/elasticidade 0,8 e aprendizado editável. Leituras GET das quatro fontes existentes, com configuração anônima pública gerada pelo mesmo verificador do console. As fórmulas financeiras e regras do indicador foram preservadas. Não grava orçamento em Ads nem altera fontes financeiras ou permissões.

Snapshots de cenários e histórico de etapas permanecem locais ao navegador/origem; cópias salvas no domínio do ChatGPT não são migradas silenciosamente para o GitHub Pages. O modelo do simulador não equivale a previsão causal ou garantia de resultado. O horizonte do plano permanece 26/10/2026–31/01/2027; a base histórica avança com os dias encerrados.

Validação local: 108 testes (96 de domínio/leitura e 12 de paridade), TypeScript e release aprovados. Navegador: abertura pela sidebar real, busca, recarregamento pelo link direto, base 24–25/10 nas quatro condições e menu móvel em 390×844 sem overflow horizontal. Fonte e saída `console/` são publicadas juntas; HTML original, backend, login e tarefas permanecem preservados.


## Análise por bases e liberação do primeiro aumento — 26/10/2026

A aba Plano até janeiro compara simultaneamente as bases após 16/10 (desde 17/10), 7, 14 e 30 dias encerrados. Mesmo horizonte 26/10/2026–31/01/2027, controle R$ 2.790/dia e calendários gradual/direto existentes. O motor único agora aceita uma janela histórica explícita; não há cópia de fórmulas. Modelos CAC constante/elasticidade 0,8 e com/sem aprendizado estão disponíveis para as quatro bases, incluindo tabela integral das 16 condições e JSON com fontes, parâmetros e resultados.

Entrada padrão: elasticidade 0,8 com aprendizado (7 dias/6% por aumento, editável). Cartões exibem lucro de manter/gradual/direto, ganhos sobre manter, CAC e lucro diário observados. O destaque de maior lucro inclui manter e empates. Resumo informa em quantas bases manter lidera/empata, com contagem explícita das bases disponíveis. Gráfico alterna lucro líquido total e ganho sobre manter; negativos são preservados. Bases inválidas exibem causa individual e não geram zeros ou janelas menores. Exportação não fica disponível sem resultado válido. As análises antigas dos dois dias recentes permanecem em expansão sob demanda.

Acima das projeções, critérios reais reutilizam a mesma decisão de ScalePlan, o histórico local da etapa e a maturação. O primeiro teste R$ 3.100 exige etapa inicial válida, janela completa, CAC até R$ 9, lucro diário e acumulado positivos, gasto dentro de ±20% do patamar e ausência de motivo de redução. Aumento posterior mantém regras de +5%, gasto efetivo maior e recuperação contra referência congelada. O quadro não libera aumento por projeção; mudar base/elasticidade/aprendizado não o altera. A primeira data possível é calculada com a alteração, janela e reserva de atualização, sem sugerir avanço automático nessa data. Etapa já avançada não é tratada como se estivesse em R$ 2.790. Fontes/armazenamento inválidos impedem a liberação.

Leitura conferida até 25/10: 9/14 dias após a alteração, CAC R$ 9,23; manter e reavaliar a partir de 31/10 se os demais critérios passarem. Com elasticidade 0,8/aprendizado, manter lidera nas bases 14/30; ganhos direto/gradual sobre manter conciliam exatamente com a análise calculada em 26/10. Quatorze/trinta dias misturam antes/depois e são identificados como sensibilidade. O carregamento cobre pelo menos 30 dias e o período desde 17/10, além da janela necessária à etapa, inclusive após novos registros.

Verificação: 113 testes (101 domínio/leitura e 12 paridade). Novos testes cobrem datas exatas, exclusão de 16/10 na base posterior, mesmo calendário/mídia, conciliação por estratégia, isolamento de bases sem dados, cobertura insuficiente, impostos/hipóteses inválidos, aprendizado zero, vencedor entre três opções, empate e perdas. As 16 condições conciliam com o snapshot da análise anterior. Navegador: controles recalculam sem alterar critérios, campo vazio remove projeções e desabilita exportação, JSON auditado, gráficos de total/ganho, expansão da base recente e 390×844 sem overflow horizontal. Nenhuma campanha, fonte financeira ou permissão foi alterada.


## Simulador simples e Financeiro — 29/10/2026

O simulador passa a abrir uma tela enxuta: verba diária, duração, adaptação e efeito no lucro comparado com manter. Três curvas em um gráfico, com pico teórico limitado, permitem entender a margem das vendas adicionais. As quatro áreas completas anteriores ficam em “Ferramentas avançadas e cenários salvos”, sem apagar armazenamento local. O indicador real permanece independente das hipóteses; pendências na DRE impedem apresentar a liberação do painel como validação do lucro completo.

Novo módulo Financeiro na navegação principal e busca (`#/financeiro`), DRE mensal com CSV, resultado parcial explícito e complementos privados para o gestor. Receita já exclui cancelados; estornos são informativos e não descontados duas vezes. Migração aditiva 021, RLS e RPC versionada foram aplicadas isoladamente, com testes transacionais e ROLLBACK de fixtures. Nenhuma despesa fictícia persistida ou campanha alterada. Método, privacidade, testes e rollback: [financeiro-dre.md](financeiro-dre.md).


## Resultados Financeiros editáveis — 29/10/2026

O simulador adota os mesmos 15 cards de Resultados Financeiros da Visão Geral, na mesma ordem, com lápis para premissas e cadeados explicativos para resultados. Receita permite escolher preço ou volume; CAC/transações/conversão usam um único driver para evitar circularidade; upsell exige receita e custo incrementais. Histórico calibra unidades e taxas, mas elasticidade não é tratada como medida causal. Controle, indicador real, DRE privada, fontes e snapshots avançados preservados.

Leituras adicionais somente agregadas: upsell pago, checkouts do projeto original e acumulado de 2026. Fonte opcional desconhecida nunca vira zero. Sem alteração no banco/Ads. 140 testes (128 domínio/leitura, 12 paridade). Fórmulas, UX e limites: [simulador-cards.md](simulador-cards.md).

## Redesign visual · 29/10/2026

Novo shell claro com rail integrado, cabeçalho em cápsulas e logo oficial sem wordmark. Visão Geral mantém 15 indicadores em 5×3, fórmulas originais e demais painéis, com estados de resultado, comparação e meta separados. Progresso mensal explicita escopo e não mistura meses. Financeiro e Simulador recebem a mesma família visual, preservando permissões e cálculos. Build/release local e 144 testes aprovados; detalhes em [redesign-console-2026-10-29.md](./redesign-console-2026-10-29.md). Sem push/publicação nesta etapa.

## Controle de usuários · 30/10/2026 (backend ativado, release GitHub Pages)

Nova entrada do console com sessão e membership ativa, perfis gerente/administrador e administração real em `#/usuarios`. Criação de contas exclusiva de administradores no servidor; troca de perfil, suspensão e reativação com versão/auditoria e proteção do último administrador. A migração 022 foi aplicada e a função `console-users` está ativa com validação JWT. Cadastro público desabilitado; conta preexistente habilitada como administrador. Testes de autorização passaram também no backend real, com ROLLBACK e conferência de ausência de fixtures. Fontes públicas e dashboard legado preservados. Instruções completas e evidências de validação: [Controle de usuários](usuarios-painel.md).

## Inicial de resumo · 30/10/2026

A inicial passa a usar composição React própria, com panorama financeiro, receita recente, meta de receita, execução e canais da operação, além dos atalhos de cada módulo. Usa os totais e o filtro salvo da Visão Geral, incluindo hoje quando selecionado; despesas complementares permanecem na DRE. Os períodos e limites das amostras são explícitos. Mantém loader global e permissões. Fonte e validação: [Inicial de resumo](inicio-resumo.md).

## Alinhamento de indicadores e metas · 30/10/2026

A inicial lê o snapshot do motor original da Visão Geral, com os mesmos quatro valores e período salvo, incluindo o dia atual. Migração 023 aplicada no servidor: apenas administrador ativo altera metas/imposto. Gerente consulta em modo de leitura. Validação: 180 testes e build aprovados, autorização também testada no banco remoto com rollback. Detalhes em [Inicial](inicio-resumo.md) e [Usuários](usuarios-painel.md).

## Filtro na inicial e edição de usuários · 30/10/2026

A inicial permite aplicar presets e datas personalizadas compartilhadas com a Visão Geral, sem navegação adicional. Administração de usuários inclui nome/email com revisão e redefinição direta de senha com confirmação. Migração 024 e função console-users versão 2 ativas: mudanças de Auth e cadastro são atômicas, com autorização e versão rechecadas no banco e auditoria sem credenciais. Validação: 184 testes da aplicação, 11 da função, SQL local/remoto e login real usando contas descartáveis removidas ao final.

## Recuperação por email · 30/10/2026

Ação administrativa de envio de link e página independente de criação de senha. Template/redirect configurados; migração 025 aplicada, signup mantido fechado. Sessão de recuperação isolada, token único/expirado, confirmação de senha e conta ativa. Envio Google pendente de autorização SMTP de suporte@verificaplaca.com.br. 188 testes app +15 Edge aprovados; SQL remoto e Auth real sem mensagens aprovados. [Ativação](email-recuperacao.md).

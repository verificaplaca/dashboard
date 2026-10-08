# Console — primeira integração de produção

Fonte canônica: `console-app/`, no repositório `verificaplaca/dashboard`.
Saída publicada: `console/`, pelo GitHub Pages existente (main, raiz).
O endereço original e `dashboard.html` permanecem disponíveis.

## Escopo entregue

- Inicial com saudação pelo horário de São Paulo, resumo financeiro real e acesso aos módulos.
- Visão Geral com o motor original integral, sem iframe. Mantém os 15 indicadores, filtros, impostos, herança de metas e comparativos.
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

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
- Demonstração completa preservada em `?demo=1`, com armazenamento e comandos fictícios separados do modo de produção.

Nenhuma migração SQL, alteração de políticas, mudança de cron, sync ou escrita em pedidos foi executada. Login, edição de perfil e de metas só ocorrem por ação do usuário autenticado. O banco mantém a regra existente: qualquer usuário `authenticated` autorizado pela RLS pode editar metas. Isso ainda não representa os papéis e equipes da demonstração.

## Contratos que faltam para a operação completa

CRM, responsabilidade por clientes, tarefas, suporte, chat compartilhado, instâncias WhatsApp, convites, RBAC por equipe e progressão individual não têm tabelas/APIs próprias no backend auditado. Permanecem em integração no modo real; o botão de demonstração permite revisar a experiência já aprovada.

Pedidos sincronizados usam `(provider, provider_order_id)`; tracking usa `checkout_id`. Não inventar vínculos entre as duas identidades. Não usar `raw_json`, emails e telefones nas leituras públicas do console. O projeto do site (`ozquoloetuzynnyzkado`) tem checkouts e consultas; o projeto analytics (`ftmgmfdqdqxboiktxcoj`) reúne financeiro e tracking. As credenciais privilegiadas do site devem permanecer no servidor.

Próxima etapa: definir provedores WhatsApp, contratos de consulta/entrega e backend de operação com autorização, persistência, auditoria e eventos reais para gamificação. Não converter dados fictícios em registros de produção.

## Preservação e verificação

Em 07/10/2026, 18:40:35 UTC, auditoria somente leitura do projeto analytics confirmou 48.306 pedidos, 48.306 itens, 18.909 dispatches e 3 meses de metas. Contagens são uma fotografia de uma base que continua recebendo syncs, não um backup restaurável.

O SHA-256 de `dashboard.html` e da referência congelada é `9d2852deed6eb03200c4836adf86bc3de1e31ab61543497a0834553674bfdedb`. O release é bloqueado se a fonte canônica divergir da referência: revisar e recapturar deliberadamente antes de atualizar o motor.

43 testes cobrem o domínio demonstrativo, a apresentação, equivalência de fontes/cálculos, validação dos valores mensais, rejeição de identidade ausente, concorrência e falha de RLS. Testes de escrita usam cliente controlado; não alteram registros reais. O login e a persistência de edição precisam de validação com a conta do operador, sem pedir senha em chat.

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

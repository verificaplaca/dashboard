# Controle de usuários do painel · 30/10/2026

Backend ativado no projeto analytics `ftmgmfdqdqxboiktxcoj` em 30/10/2026: migração 022 aplicada isoladamente, função `console-users` ativa com validação JWT e cadastro público de Auth desativado. A conta preexistente recebeu o perfil administrador. Nenhum usuário real novo foi criado, nenhuma conta real foi suspensa e nenhuma credencial privilegiada foi incluída no frontend. A interface está publicada em https://verifica-placa-ops.vercel.app/#/usuarios pela integração Vercel, a partir do commit `d69e243`. O mesmo release `console/` acompanha a publicação secundária GitHub Pages.

## Perfis

| Perfil | Consultar painel e metas | Alterar metas e imposto | Criar usuários | Gerenciar perfis e acesso |
|---|---|---|---|---|
| Gerente | Sim | Não | Não | Não |
| Administrador ativo | Sim | Sim | Sim | Sim |

Rota real `#/usuarios`, exibida apenas para administradores. Busca por nome, paginação de 50 registros e formulário com nome, email, senha inicial (12–128 caracteres) e perfil. A senha é enviada por HTTPS ao servidor, sem persistência pela aplicação, logs ou auditoria. A criação não envia emails. O administrador deve entregar as credenciais por canal privado. O botão Redefinir senha permite ao administrador definir e confirmar uma nova senha de 12–128 caracteres, que substitui a atual imediatamente. A senha deve ser entregue ao usuário por canal privado; o envio de link de recuperação também foi implementado, mas depende da ativação do remetente Google descrita em `email-recuperacao.md`. Gerenciar permite editar nome, email de acesso, perfil e situação, com uma etapa de revisão. O novo email passa a valer imediatamente para login. Convites com aceite não fazem parte desta entrega.

## Autorização

`console_users` guarda os dois perfis e a situação de acesso. Metadata editável de Auth não concede permissões. O painel exige sessão e membership ativa antes de montar seus módulos. A identidade e a membership são consultadas novamente a cada 30 segundos enquanto a janela está visível e ao voltar à janela. Falha de autorização bloqueia o painel; retorno de dados anteriormente lidos não é apagado do dispositivo.

A Edge Function verifica o JWT pelo Auth e consulta o perfil do chamador antes de criar uma conta. A chave administrativa permanece exclusivamente no ambiente do servidor. A RPC `console_provision_user` só é executável pelo servidor e verifica novamente o administrador na transação que registra membership e auditoria. Falha de provisionamento tenta remover exclusivamente a conta Auth recém-criada; falha de compensação registra apenas o UUID para manutenção. Em resultado de rede incerto, atualizar a lista antes de repetir.

Alterações de perfil/situação passam pela RPC autenticada `console_update_user`, com versão observada e auditoria. Um lock transacional serializa alterações e protege o último administrador ativo, inclusive contra alterações concorrentes. O usuário não pode escrever diretamente na tabela ou apagar a auditoria. Não existe exclusão definitiva pela interface.

Memberships do console sincronizam as memberships preexistentes de tarefas. Ambos os perfis têm o acesso operacional de gestor. Suspensão revoga também tarefas e despesas privadas. A migração 022 passou a exigir membership ativa para escrita em metas; a migração 023 restringe INSERT, UPDATE e DELETE a administradores ativos. Gerentes mantêm leitura. O salvamento verifica novamente identidade e perfil no banco, além dos campos bloqueados e botão de salvar ausente para gerentes. Acesso privado não depende somente da ocultação dos botões.

**Limite de escopo:** os agregados financeiros e endpoints de leitura que já eram públicos continuam públicos, assim como `dashboard.html` e a demonstração isolada `?demo=1`. O login protege a entrada do novo console; não transforma o dashboard legado nem suas fontes públicas em dados confidenciais. Uma privatização integral dessas fontes exigiria migrar os consumidores existentes para consultas autenticadas.

## Ativação

Publicar backend antes do novo frontend, para evitar bloquear a conta atual por ausência da tabela:

1. Manter o cadastro público de Auth desativado (`Allow new users to sign up` desligado).
2. Aplicar **somente** `supabase/migrations/022_console_users.sql` no projeto analytics `ftmgmfdqdqxboiktxcoj`; não executar todas as migrações legadas com `db push`. O bootstrap exige exatamente uma conta Auth não excluída e uma membership ativa de gestor correspondente. A conta existente vira administrador; novas contas não são criadas automaticamente. Se a identidade não corresponder ou houver múltiplas contas, a transação aborta para revisão explícita.
3. Publicar `supabase/functions/console-users/index.ts`, incluindo `validation.ts`, mantendo a validação JWT da plataforma habilitada. Usa somente `SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY` já fornecidos pelo ambiente Supabase.
4. Publicar fonte e pasta gerada `console/` juntas pelo fluxo GitHub Pages existente.
5. Entrar com a conta atual e validar a tela Usuários. Criar contas reais apenas por solicitação do operador. Não colocar senhas em chat, arquivos ou evidências.

## Verificação

- 155 testes de domínio/API e 12 testes de paridade do motor original aprovados.
- Seis testes da Edge Function: acesso anônimo, gerente/admin inativo, campos forjados, identidade verificada, compensação e token inválido.
- PostgreSQL 17 local: migrações 013/017/018/021/022 e testes transacionais de usuários, tarefas e financeiro aprovados. Todos os fixtures de testes revertidos. A emulação local utiliza tabela Auth mínima; A migração também foi aplicada e os testes de usuários passaram no Supabase real, com ROLLBACK. A conferência posterior confirmou uma conta, um administrador ativo na versão 1, zero eventos e zero contas de teste.
- Navegador isolado com respostas fictícias: login, lista administrativa, criação de gerente, suspensão, último administrador protegido, rota/menu restritos ao administrador, bloqueio de conta suspensa e tela de 390×844 sem overflow horizontal da página. Nenhuma sessão de usuário real foi gerada para estes testes.
- Verificação online em Vercel: entrada exige login; bundles da rota `/usuarios` disponíveis; diretório de usuários anônimo rejeitado com HTTP 401. Publicação Vercel confirmou sucesso no commit `d69e243`.
- Função remota: status ACTIVE, verify_jwt=true; chamada sem sessão de usuário rejeitada com HTTP 401. Cadastro público confirmado desativado.
- TypeScript, build e release local aprovados. O script de cálculo original continua íntegro.

Executar testes de backend com `deno test supabase/functions/console-users/index.test.ts`. Executar `supabase/tests/console_users.sql` em banco de teste após a migração, com o único administrador de bootstrap; a suíte usa fixtures transacionais e ROLLBACK. O teste antigo de tarefas também continua exercitando o papel legado de membro com seu escopo restrito.

## Reversão

Reverter fonte/release do frontend não deve apagar contas, memberships, tarefas, despesas nem auditoria. Manter as tabelas e a restrição de escrita de metas. Se o frontend antigo for restaurado, ele volta a ter entrada pública; não apresentar esse rollback como preservação da exigência de login do novo console. Retirar a Edge Function impede novas criações pela tela sem eliminar contas já provisionadas.

## Metas exclusivas de administradores · 30/10/2026

Migração 023 aplicada isoladamente no servidor. `supabase/tests/targets_administrator_only.sql` validado localmente e no projeto remoto, com rollback de todas as fixtures. Metas e contas reais preservadas. Gerentes consultam mês e histórico sem edição. Metadata de Auth não concede escrita.

## Edição de dados e redefinição de senha · 30/10/2026

Migração 024 e Edge Function `console-users` versão 2 publicadas. A função verifica a identidade do chamador e permite as ações `update` e `reset_password` somente para administradores ativos. O alvo e a versão observada são validados; nome/email são normalizados e campos de autoridade forjados são rejeitados.

A RPC exclusiva do servidor `console_prepare_user_change` cria uma intenção de dois minutos em `console_private.user_changes`, sem senha ou hash. O Admin API oficial altera Auth usando um marcador em app metadata. O trigger `console_finish_user_change` revalida o administrador e a versão sob o mesmo lock das alterações de perfis, sincroniza console/tarefas, audita a ação, consome a intenção e remove o marcador. Tudo ocorre na transação do Auth: falha ou revogação reverte também email/senha, evitando cadastros divergentes. A proteção do último administrador permanece. Falhas retornadas pelo Auth cancelam a intenção; intenções expiradas são eliminadas na próxima preparação.

Auditoria não guarda senha, hash ou marcador. O novo email é confirmado administrativamente; a revisão informa que ele substitui o login atual. A redefinição não tem duração temporária nem força troca no próximo acesso. Sessões seguem o comportamento do Supabase Auth ao alterar a senha.

Validação: 11 testes da Edge Function, testes de validação/cliente na aplicação e `supabase/tests/console_user_details.sql` local/remoto com rollback. Teste completo da função publicada com contas descartáveis confirmou login pelo novo email, rejeição do antigo, troca de senha, rejeição da senha anterior e bloqueio do gerente. Todas as contas descartáveis e suas memberships/auditorias foram removidas; contas reais preservadas.

## Link de redefinição por email

Administradores podem solicitar link ao destinatário cadastrado e ativo. Migração 025 e função atualizada adicionam autorização, controle de reenvio e auditoria. A página `?recovery=1` valida token único em sessão isolada e permite ao usuário escolher sua senha. Remetente escolhido: suporte@verificaplaca.com.br pelo Google Workspace. Envio efetivo bloqueado até o dono autorizar SMTP; nenhuma mensagem real enviada nos testes. Ativação: [Email de recuperação](email-recuperacao.md).

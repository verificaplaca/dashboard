# Redefinição de senha por email

A ação **Enviar link por email** fica em Usuários, somente para administradores ativos. Confirma o destinatário, consulta seu email e sua versão no banco e exige conta ativa. Há cooldown de 60 segundos e limite de cinco tentativas por usuário/hora, além dos limites do provedor. O envio não muda a senha atual. As tentativas ficam em tabela privada; apenas envios aceitos são registrados na auditoria administrativa, sem tokens ou credenciais.

## Estado atual

Migração 025 e função atualizada publicadas. Template de recuperação em português, site URL de produção e redirect permitido `https://verifica-placa-ops.vercel.app/?recovery=1` configurados no Supabase. Cadastro público permanece fechado.

**O envio efetivo está desativado até autorizar o Google Workspace.** A caixa `suporte@verificaplaca.com.br` já existe no Gmail e o usuário concluiu a criação da senha de app; falta inseri-la no ativador para configurar SMTP. O servidor exige `CONSOLE_RECOVERY_EMAIL_ENABLED=true`; sem isso informa a pendência e não anuncia sucesso. O SMTP de testes do Supabase não é usado como alternativa silenciosa.

## Link e nova senha

O template leva `TokenHash` no fragmento da URL, que não é enviado ao servidor web e é removido imediatamente pelo bootstrap. O usuário clica em Continuar para validar o link; só esse clique troca o token pelo Auth. Uso único e expiração são verificados no Supabase.

A página usa cliente Auth isolado, sem persistência e sem detecção automática de sessão da URL. Não usa nem substitui uma sessão administrativa já aberta. Após validar identidade e membership ativa, mostra nova senha e confirmação (12–128 caracteres), revalida o acesso antes de salvar e encerra a sessão isolada após sucesso. Não monta o painel normal. Link inválido, expirado ou já consumido não libera o formulário. Recarregar após a remoção do fragmento exige reabrir o link do email ou pedir outro se já consumido. O fluxo implícito legado com `type=recovery` também é reconhecido em cliente isolado.

A redefinição direta pelo administrador permanece. A conclusão pelo próprio usuário usa o Auth oficial e seu registro de segurança; a auditoria de console desta entrega registra o pedido de email, não a senha escolhida.

## Ativação pelo dono da conta Google

1. Confirme que o remetente é uma caixa Google Workspace existente. Entre na conta **suporte@verificaplaca.com.br**, habilite verificação em duas etapas e abra [Senhas de app](https://myaccount.google.com/apppasswords). Crie uma senha de app para “Verifica Placa”. A organização pode precisar liberar essa opção. Não use a senha principal nem compartilhe a senha aqui no chat. [Requisitos oficiais do Google](https://support.google.com/accounts/answer/185833?hl=pt-BR).
2. Preferência: abrir a página local de ativação, colar a senha no campo oculto e clicar em **Autorizar e ativar envio**. A página é iniciada por `scripts/google-recovery-setup.py`, aceita conexões somente em 127.0.0.1, tem endereço aleatório e valida origem/CSRF. Não registra a senha, não usa recursos externos e encerra após sucesso ou 15 minutos. A criação e a entrada da credencial são feitas pelo usuário.

   Alternativa no Terminal local:

   ```sh
   python3 /Users/danielmacedo/Documents/ChatGPT/verificaplaca/dashboard-published/scripts/activate-google-recovery-email.py
   ```

3. Cole a senha de app apenas no campo oculto do Terminal. O script valida o SMTP com o Google sem enviar mensagem, configura o remetente no Supabase e ativa o envio. Requer Supabase CLI autenticado. No Mac lê a credencial existente no Chaves, em memória. Não coloca segredos em argumentos, arquivos, saída ou Git. Antes de anunciar sucesso, relê a configuração no projeto e confirma o valor exato do flag de ativação pelo digest do segredo.

Parâmetros: `smtp.gmail.com`, SSL porta `465`, usuário e remetente `suporte@verificaplaca.com.br`, nome “Verifica Placa”. [Configuração oficial de aplicativos no Workspace](https://knowledge.workspace.google.com/admin/gmail/send-email-from-a-printer-scanner-or-app).

Alternativa manual: configurar esses campos em [SMTP do projeto](https://supabase.com/dashboard/project/ftmgmfdqdqxboiktxcoj/auth/smtp) e ativar `CONSOLE_RECOVERY_EMAIL_ENABLED=true` nos segredos da função. A criação e entrada de uma nova credencial no navegador precisam ser realizadas pelo usuário.

## Validação

188 testes da aplicação (176 domínio +12 paridade) e 15 da Edge Function aprovados. SQL local/remoto cobre autorização, versão, conta suspensa, destinatário, cooldown e auditoria idempotente, com ROLLBACK. Teste real do Auth com conta descartável e `generate_link` (sem email) confirmou usuário correto, bloqueio de reutilização, nova senha válida e rejeição da antiga. Conta e registros de teste removidos. Entrega em caixa postal ainda não testada: falta a autorização Google e nenhum email foi enviado a usuários reais.

Verificação posterior à criação da senha de app: Google SMTP, remetente e flag ainda ausentes no servidor. A mensagem relatada pelo usuário correspondia à conclusão da etapa no Google, não à execução do ativador. O ativador foi reforçado com pós-condições e uma tela local sem terminal. Oito testes adicionais cobrem persistência/flag, erro de login Google, origem/CSRF, ausência de eco de credenciais e sucesso.

Correção do formulário local: o Chrome envia `Origin: null` em POST de navegação com `Referrer-Policy: no-referrer`. A página agora usa `same-origin`, mantendo o endereço privado fora de referências externas e preservando a origem esperada no envio. Reproduzido e corrigido no Chrome com formulário sem credenciais. Validação CSRF, Host e origem exata mantidas; o teste continua rejeitando origem `null` ou externa. Oito testes Python aprovados. Reabrir a página corrigida e inserir pessoalmente a mesma senha de app.

Correção adicional: a Management API exige `smtp_port` como string. O ativador enviava número e a API retornava HTTP400; o formulário mostrava incorretamente um aviso genérico sobre a senha. Confirmado no endpoint real sem credenciais SMTP: número400, string200. Porta corrigida para string, persistência conferida com até oito leituras e mensagens/registro seguro por etapa. Conexão TLS real com Gmail465 e587 aprovada; formulário completo aprovado no Chrome usando operação fictícia. Servidor local passou a atender conexões simultâneas de navegador, serializando a ativação. Doze testes Python aprovados. Reutilizar a senha de app já criada; não a solicitar no chat.

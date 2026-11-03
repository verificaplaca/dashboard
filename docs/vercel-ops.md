# Hospedagem ops.verificaplaca.com.br

Endereço principal: https://ops.verificaplaca.com.br/. O endereço verifica-placa-ops.vercel.app redireciona permanentemente (308) ao domínio oficial, preservando caminhos, parâmetros e fragmentos do navegador. Configuração atualizada em 03/11/2026. Os registros de implantação abaixo mantêm o histórico da migração.

Destino solicitado em 30/10/2026: Vercel, projeto `verifica-placa-ops`, repositório `verificaplaca/dashboard`, branch `main`.

## Configuração

Importar o repositório existente com Root Directory `.` (raiz), preset Vite e configurações de `vercel.json`. A instalação usa `npm ci --prefix console-app`; o build usa `npm run build:vercel --prefix console-app`; somente `console-app/dist` é servido. Não selecionar `console-app` como Root Directory: o gerador lê `dashboard.html` da raiz.

Entrada do console em `/`, com navegação por hash (`/#/inicio`, `/#/dashboard`, `/#/simulador`, `/#/financeiro`). O dashboard original e seus painéis continuam disponíveis em `/dashboard.html`, `/google-ads-cac.html` e `/tracking-gateway.html`, copiados por lista explícita sem mudar seu conteúdo. Os caminhos antigos `/console/` e `/dashboard/console/` redirecionam à raiz. GitHub Pages continua disponível para retorno durante a migração.

Nenhum segredo de servidor é necessário para o frontend. O build verifica as credenciais anônimas públicas e o papel `anon` dos dois projetos Supabase existentes. Não enviar `service_role`, credenciais Pagar.me, tokens Ads ou variáveis do backfill à aplicação. Migrações, sincronizações e funções Supabase existentes permanecem em seus ambientes atuais; publicar o frontend não executa esses processos.

## Domínio

Adicionar `ops.verificaplaca.com.br` ao projeto Vercel e usar exatamente o destino DNS informado por ele. O domínio usa nameservers Cloudflare. Criar apenas o registro do subdomínio `ops`, inicialmente DNS only, sem alterar os registros do site ou email. Conferir domínio válido, certificado HTTPS e leitura dos dados após a propagação.

## Login e dados locais

O login usa email/senha do Supabase existente, sem callback OAuth ou magic link no fluxo atual. O operador precisa entrar novamente na nova origem. Permissões e dados privados continuam controlados pelo backend/RLS. Não transferir tokens entre origens.

Cenários salvos, histórico de etapas e preferências em armazenamento local pertencem à origem antiga e não aparecem automaticamente no subdomínio. Preservar o endereço anterior para exportar/reimportar cenários pelas ferramentas existentes; não apagar o armazenamento antigo nem reiniciar etapas reais sem revisar esse histórico.

Hospedar o sistema não conclui módulos operacionais ainda marcados como demonstração/integração. Mantêm-se os mesmos módulos e limites documentados em console-producao.md.

## Validação e retorno

Executar `npm test`, `npm run release` (confere a publicação existente) e `npm run build:vercel` em `console-app/`. Verificar a saída explícita, páginas legadas, arquivos estáticos, Visão Geral com dados reais, Financeiro, Simulador, navegação e login. O build Vercel aponta o link Dashboard original ao mesmo domínio; o build Pages mantém o link anterior.

Rollback do frontend: restaurar o deployment anterior na Vercel ou usar temporariamente o endereço GitHub Pages. Não remover tabelas, despesas, tarefas ou histórico do Supabase. Reverter hospedagem não deve reverter escritas do operador.

## Estado

Publicado em produção em 30/10/2026: https://verifica-placa-ops.vercel.app/.

- Projeto: https://vercel.com/verificaplacas-projects/verifica-placa-ops
- Deployment: `dpl_2PrvygTTCecR8WN9guLYqcjhrDex`
- Fonte: `7d2c41879147dcf44a3d5b536c6c21dbeb0597ed`, branch main, publicação automática pelo GitHub.
- 140 testes aprovados; build Pages e Vercel aprovados. CI Console checks e Pages no GitHub concluídos com sucesso.
- HTTP 200 no console, modelo da Visão Geral e páginas/serviço legados. HTML legado idêntico à fonte. Redirecionamentos dos caminhos antigos e cabeçalho noindex conferidos em produção.
- `ops.verificaplaca.com.br` adicionado ao ambiente Production, pendente de DNS e verificação de propriedade. A Vercel informa que o domínio já está vinculado a outra conta Vercel; nenhum projeto ou domínio existente foi removido ou transferido.
- Usuário confirmou que perdeu o acesso Cloudflare e pedirá nova liberação ao desenvolvedor. Não houve alteração no DNS.

Registros exatos exibidos pela Vercel em 30/10/2026, na zona `verificaplaca.com.br`:

| Tipo | Nome | Conteúdo | Proxy |
| --- | --- | --- | --- |
| CNAME | ops | 91fbe417f5c8ecdf.vercel-dns-017.com. | Desativado (DNS only) |
| TXT | _vercel | vc-domain-verify=ops.verificaplaca.com.br,614835aed71b436cf4d7 | DNS only |

Adicionar o TXT de verificação sem apagar outros valores TXT em `_vercel`. O TXT acima pode ser removido após a verificação ser concluída. Usar apenas o subdomínio ops; preservar o domínio principal, MX e demais serviços. Após a liberação, conferir DNS na Vercel, certificado HTTPS, navegação, dados reais e login no endereço final. Até lá, usar o endereço provisório acima.


Validação no navegador publicado: Visão Geral carregou os 15 indicadores e as fontes reais; Financeiro carregou receita de R$ 136.412,32 e 8.442 pedidos até 29/10; Simulador carregou os 15 cards e a base histórica até 29/10. Nenhum erro de execução observado. Formulário de login existente abre corretamente; salvamento autenticado não foi executado sem o login do operador. Banco, permissões e integrações existentes preservados.

## Domínio validado · 03/11/2026

Após o desenvolvedor configurar os registros no Cloudflare, CNAME e TXT foram confirmados no DNS público (Cloudflare 1.1.1.1 e Google 8.8.8.8). A ação Refresh em Project Domains concluiu a verificação na Vercel: `ops.verificaplaca.com.br — Valid Configuration — Production`.

HTTPS validado sem ignorar certificados. O HTML do domínio oficial é idêntico ao de produção em verifica-placa-ops.vercel.app. Modelo da Visão Geral, dashboard legado e robots.txt responderam 200. O Chrome abriu o domínio normalmente e exibiu a tela de login existente; é necessário entrar novamente por ser uma origem diferente. Nenhuma sessão foi transferida e nenhum registro do domínio principal ou email foi alterado.

O resolvedor do terminal ainda apresentou cache negativo durante parte da conferência; os testes HTTP nesse período usaram o endereço A publicado pelo DNS, preservando o hostname TLS e a validação do certificado. A abertura normal no Chrome confirmou acesso pelo nome.

Evidências locais: `output/dominio-ops-2026-11-03/vercel-validado.png`, `login-dominio-oficial.png`, `http-validacao.json`.

## Endereço principal e links ativos · 03/11/2026

O domínio Vercel padrão foi alterado em Project Domains para `Redirect to Another Domain → 308 Permanent Redirect → ops.verificaplaca.com.br`. Conferidos HTTP 308 e Location com caminho e query preservados; navegador confirmou também `#/simulador` preservado. Domínio oficial permanece conectado a Production.

Supabase Auth: site_url alterado para `https://ops.verificaplaca.com.br/`; redirect permitido adicionado para `https://ops.verificaplaca.com.br/?recovery=1`. A permissão anterior foi mantida para links de recuperação já emitidos, que agora redirecionam ao domínio oficial. Template usa RedirectTo, sem hostname fixo. SMTP, cadastro e autenticação preservados.

A função console-users passou a gerar recuperação no domínio oficial e foi implantada no projeto ftmgmfdqdqxboiktxcoj. Fonte remota baixada para diretório temporário conferida após a implantação. Quinze testes da função aprovados. Nenhum email foi enviado nem senha alterada na validação. Documentação ativa de produção, usuários, recuperação e CONTEXT atualizada; evidências datadas e URLs históricas permanecem como registros do que foi executado.

Evidência da configuração: `output/dominio-ops-2026-11-03/dominio-principal.png`.

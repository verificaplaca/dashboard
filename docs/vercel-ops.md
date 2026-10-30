# Hospedagem ops.verificaplaca.com.br

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

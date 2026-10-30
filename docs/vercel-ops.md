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

Preparação local concluída; implantação e validação do DNS ainda pendentes. Atualizar esta seção com o deployment e a verificação final.

# Desativação do painel legado · 04/11/2026

Painel principal: https://ops.verificaplaca.com.br/.

O novo console usa `overview/model.json` local, gerado da referência preservada, e os endpoints Supabase existentes. Não depende da disponibilidade HTTP de `dashboard.html`, `google-ads-cac.html` ou `tracking-gateway.html`. As rotinas de sincronização GitHub Actions chamam diretamente as Edge Functions Supabase e continuam habilitadas.

A publicação Vercel deixou de copiar as três páginas antigas e o serviço JavaScript de CAC. A página independente `overview/index.html` também foi removida do pacote. Seus endereços redirecionam permanentemente para os módulos do novo painel. Imagens compartilhadas, modelo nativo e fontes de dados foram preservados. O domínio padrão da Vercel já redireciona ao domínio oficial.

A hospedagem GitHub Pages do repositório é desativada, incluindo a cópia secundária do console. O repositório e os arquivos de referência permanecem disponíveis para desenvolvimento, testes de paridade e recuperação. O arquivo raiz `index.html` aponta ao domínio oficial caso uma hospedagem seja reativada no futuro.

Validação: testes de domínio e paridade, build Vercel, ausência das páginas legadas no pacote, redirecionamentos HTTP, modelo nativo disponível e desativação confirmada pela API do GitHub Pages.

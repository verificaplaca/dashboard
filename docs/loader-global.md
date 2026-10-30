# Loader global do painel

O componente fornecido foi integrado em `console-app/src/components/ui/loader-10.tsx`. Mantém o efeito gooey, as duas animações, as cores configuráveis e `forwardRef`. Recebe IDs de filtro únicos por instância, composição do `style` externo, rótulo em português e versão estática para quem prefere movimento reduzido. No painel, as cores seguem o verde da interface.

`LoadingProvider` compartilha o estado entre as leituras da tela. Cada leitura possui uma inscrição própria: terminar uma requisição não libera o painel enquanto outra estiver pendente. A saída aguarda 160 ms para que os dados e leituras encadeadas assentem, seguida de um fade de 320 ms que revela o painel pronto. A camada só é removida após a transição; foco e interação continuam protegidos nesse intervalo. Uma nova leitura cancela a saída e restaura a cobertura imediatamente. A preferência por movimento reduzido pula o fade. O espaço da barra de rolagem é reservado para evitar deslocamento lateral ao liberar a página. Limpeza ao desmontar/cancelar libera somente a inscrição correspondente. O conteúdo permanece montado para que a carga termine; durante a espera, uma camada opaca cobre toda a janela, bloqueia interação, protege o foco e fornece um único anúncio acessível. Não há dados parciais visíveis atrás do loader.

Cobertura: entrada dos módulos e restauração de acesso, Início/resumo executivo e tarefas, Visão Geral, Campanhas, Financeiro/despesas, Simulador/histórico/áreas avançadas, Metas, Tracking, Usuários, Tarefas e histórico de tarefas. Os estados locais continuam disponíveis para tratamento de erros e desabilitação de controles. As atualizações silenciosas que já existiam (sessão, quadro e resumo de tarefas) continuam silenciosas; a consulta periódica de histórico já carregado também não interrompe o operador. O loader não simula porcentagem e erros continuam visíveis depois da espera.

Na integração do motor original, as leituras são acompanhadas até a transferência do corpo da resposta. A camada global substitui o indicador local sem alterar o script de cálculo congelado nem a função global `fetch`. Cada leitura nativa recebe limite de 30 segundos, permitindo os estados de erro/fallback já existentes. O dashboard HTML legado é preservado.

## Estrutura de componentes e estilos

O projeto já era React/TypeScript/Vite, com componentes compartilhados em `src/components.tsx` e estilos em `src/styles.css`, `src/layout.css` e CSS dos módulos. Os componentes reutilizáveis do shadcn agora têm pasta própria em **`src/components/ui/`**, resolvida pelo alias `@/components/ui`. Essa pasta evita misturar componentes de domínio com peças visuais reutilizáveis e permite que o CLI do shadcn instale novos componentes no lugar esperado. Os componentes anteriores continuam com suas importações.

- `src/lib/utils.ts`: `cn`, com `clsx` e `tailwind-merge`.
- `src/tailwind.css`: tema e utilitários Tailwind 4, sem Preflight; o reset e os tokens existentes continuam a controlar o painel.
- `components.json`: configuração shadcn para React/TypeScript, aliases e arquivo Tailwind.
- `vite.config.ts` e `tsconfig.json`: alias `@/*` para `src/*`.
- `src/loading.css`: apresentação da camada global; animações do componente seguem junto ao componente enviado.
- `src/components/ui/demo.tsx`: exemplo fornecido, com as cores vermelhas, disponível para reutilização e fora da navegação de produção.

A configuração foi feita diretamente no projeto existente. Para reproduzir a base num projeto Vite/TypeScript novo, instalar Tailwind pelo plugin oficial (`npm install -D tailwindcss @tailwindcss/vite`), habilitar o plugin no Vite e configurar o alias `@/*`. Em seguida, executar `npx shadcn@latest init` nesse projeto e escolher React/TypeScript, `src/components/ui` e o arquivo CSS desejado. Não é necessário reinicializar o projeto atual. Aqui o Preflight foi omitido deliberadamente para preservar os estilos já existentes.

Referências oficiais: [Tailwind com Vite](https://tailwindcss.com/docs/installation/using-vite), [Preflight e instalação sem reset](https://tailwindcss.com/docs/preflight), [shadcn no Vite](https://ui.shadcn.com/docs/installation/vite) e [configuração de aliases/pastas](https://ui.shadcn.com/docs/components-json).

## Validação

Testes de controle cobrem leituras concorrentes, cancelamento/desmontagem, limpeza repetida, falha e nova tentativa, e remoção de listeners. Os testes existentes preservam cálculos e fontes originais. A validação visual usa somente uma sessão fictícia em página temporária local; nenhuma conta, permissão ou despesa real é alterada.

Validação final: 163 testes de domínio/componentes/controle e 12 de paridade aprovados. Build TypeScript/Vite aprovado. Navegador na versão compilada: carga lenta, uma única camada global, bloqueio do fundo, 15 indicadores prontos na Visão Geral, atualização, troca de rota com cancelamento, troca de mês na DRE, erro HTTP 503 com liberação da tela, formulário de novo usuário habilitado e temas claro/escuro em 390×844 sem overflow da página. Nenhuma escrita real foi executada na validação.

# Verifica Placa · Business Console

Console administrativo em React e TypeScript. A entrada padrão usa financeiro, campanhas, tracking, metas e identidade do Supabase existente. Metas e perfil exigem a conta autenticada do operador. A demonstração completa permanece separada em `?demo=1`; CRM, chat, WhatsApp, tarefas e gamificação ainda precisam de backend operacional.

Fonte canônica e publicação GitHub Pages: veja [`../docs/console-producao.md`](../docs/console-producao.md). O dashboard original permanece disponível.

## Executar

```sh
npm install
npm run dev
npm run release
npm test
```

## Estrutura

- `src/domain/engine.ts`: base determinística, comandos, permissões e métricas em centavos.
- `src/domain/engine.test.ts`: testes de conciliação, deduplicação, escopo, créditos e transições.
- `src/context.tsx`: persistência local, sessão de demonstração e navegação por hash.
- `src/components.tsx`, `src/styles.css` e `src/layout.css`: componentes, tokens visuais, temas, escala de leitura e responsividade. A camada de layout final mantém o ritmo comum entre os módulos.
- `Home`, `ExecutiveSummary` e `domain/home.ts`: inicial por horário de São Paulo, foco pessoal, prioridades, agenda e evolução conforme permissões; resumo financeiro pelo mesmo motor original.
- `Dashboard`: integra a Visão Geral original ao console, com DOM no mesmo documento, tema, navegação fixa de seções, período visível, tabela dos gráficos e exportação.
- `reference/dashboard-original.html` e `reference/provenance.json`: fonte congelada em 07/10/2026 e hash SHA-256.
- `scripts/build-overview.mjs`: adapta apenas apresentação/navegação; preserva integralmente o script de cálculo e carga.
- `OriginalOverview`: integra o motor congelado no mesmo documento, com CSS encapsulado, eventos/filtros nativos, cleanup de gráficos/listeners e carregamento das bibliotecas por demanda. Não usa iframe nem postMessage.
- `public/overview/model.json`: DOM/CSS e script integral gerados em dev/build. `index.html`, bridge e guard legados são mantidos como referência de compatibilidade; o console usa apenas o modelo nativo.
- `src/domain/overview.ts`: leitor próprio das fontes HTTPS autorizadas; não altera fetch global.
- `src/motion.ts`: entradas discretas de cards na rolagem, uma vez por elemento, respeitando movimento reduzido.
- `scripts/overview-parity.test.mjs`: testes de equivalência de fórmulas/fontes/filtros/metas, preservação dos 15 KPIs e bloqueio de escrita.
- `Commerce`, `Inbox`, `Operations` e `Team`: telas demonstrativas dos demais domínios.

## Revisar os principais fluxos

1. Início → selecionar foco → abrir prioridade/agenda → consultar metas; saudação muda por horário e exclui créditos pendentes da progressão.
2. CRM → oportunidade → gerar pedido com placa → pagamento → consulta → entrega.
3. Pedido pago → adicional → estorno → créditos revertidos e indicadores recalculados.
4. Atendimento → resposta/nota interna → tarefa, oportunidade, pedido ou chamado no contexto.
5. WhatsApp → reconectar instância alternativa → fila pendente reconciliada uma vez.
6. Administração → convite → aceite simulado → permissões → simulação de perfil.
7. Metas e evolução → desafios/conquistas → extrato pendente, confirmado e revertido.
8. Perfil → identidade/foto/preferências → recarregar para conferir persistência.
9. Administração → configurações → exportar base → restaurar com backup local.

No modo `?demo=1`, use o menu superior para simular proprietário, gestor, comercial, atendimento, marketing e financeiro. Esse seletor não autentica usuários reais. Permissões são verificadas nas ações da demonstração; em produção a API deve aplicá-las independentemente da interface.

## Dados e limites

Visão Geral: 15 KPIs, sparklines, comparativos anteriores, metas e impostos mensais, produtos, saldo real Ads, bureau e fornecedores, estornos, saúde, funil, gráficos e médias por dia da semana. Lê as 12 fontes originais via GET, com a mesma paginação e regras de fallback. Falha principal segue o original e mostra dados de exemplo com aviso explícito; não são apresentados como reais. A seleção de campanha mantém os totais globais como no original. Atualizações futuras no código da referência exigem captura deliberada, revisão e testes de equivalência; não há atualização automática do motor.

Módulos demonstrativos (`?demo=1`): referência fixa em 6 de outubro de 2026, 12h BRT. Valores em centavos; tributo hipotético de 8%. A autenticação do modo real usa Supabase Auth. Presença simultânea, papéis de equipe e os comandos operacionais demonstrativos ainda exigem backend. Os pedidos da demonstração não alimentam o dashboard original.

Pedidos pagos incluem registros posteriormente estornados. A receita bruta mantém essa história; receita após estornos e contribuição descontam devoluções. Créditos novos aguardam 7 dias e validação; uma tarefa ou mensagem não gera pontos. Eventos de entrega reconhecem o operador; venda assistida reconhece o vendedor; venda automática não cria crédito individual.

O dashboard original permanece disponível e não é alterado pelo console. Sua fonte e cálculos são reutilizados no mesmo documento da aplicação, com CSS encapsulado; não há iframe nem cópia persistida dos agregados financeiros. Os módulos demonstrativos permanecem separados dessa leitura. Antes da migração: inventariar tabelas/fontes, backup com restauração testada, comparar registros/IDs/relações/métricas, documentar diferenças de fórmulas e ensaiar rollback incluindo escritas posteriores.

Continuidade da integração no repositório: `../docs/console-producao.md`. A demonstração não migra seus registros fictícios para produção.

Revisão de UX em 07/10/2026: 16 rotas incluindo `/inicio` como entrada. Paleta azul da marca, hierarquia, respiros, contraste e layout móvel revisados. 26 testes de domínio/inicial e 12 de equivalência/apresentação. Falhas originais preservam o banner vermelho de exemplo; a inicial também sinaliza a indisponibilidade no título da fonte. A agenda usa o dia atual de São Paulo contra os registros fictícios de outubro.

Revisão de integração: removidos todos os iframes do DOM do console. Filtros, tooltips, atualização, temas, período e exportação operam diretamente na página. Chart.js 4.4.1 e datalabels 2.2.0 permanecem nas mesmas versões do original e são carregados pelo bundle sob demanda. Cards entram com 420 ms e os gráficos são reanimados uma vez ao entrar em tela, sem movimento quando a preferência do sistema pede redução. O ícone fornecido foi copiado integralmente para favicon, touch icon e marca no menu.

Revisão final de layout: escalas de 12–13 px para conteúdo e 10–11 px para metadados, controles de 40–44 px, cards com respiros comuns e tabelas com rolagem horizontal local. Dialogs têm título fixo, formulários com rolagem própria e ações visíveis; drawers mantêm cabeçalho acessível. Escape fecha somente a janela ativa, Tab circula dentro dela e o foco retorna à origem, preservando o bloqueio da página ao fechar um popup sobre um drawer. Menu de perfil fecha por Escape ou clique externo. O histórico de atendimento rola dentro da conversa sem usar scrollIntoView na página.

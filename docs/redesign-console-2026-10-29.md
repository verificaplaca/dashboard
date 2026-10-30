# Redesign do console · 29/10/2026

Implementação baseada na composição aprovada, com sidebar integrada ao mesmo fundo, navegação superior em cápsulas, logo oficial sem wordmark, respiro e tags compactas.

## Entrega

- `console-app/src/redesign.css`: shell, rail expansível de ícones, cabeçalho, aparência comum, tema claro inicial (preserva preferência anterior), responsividade e acabamento de Financeiro/Simulador.
- `console-app/public/verifica-placa-logo.png`: cópia binária do logotipo branco fornecido pelo usuário; SHA-256 `9a623373efc384890ced085350838bdb3b6846c3d8907e3077e8d341c51c43e5`.
- `public/overview/overview.css`: Visão Geral com quinze cards de iguais dimensões numa grade 5×3. Em telas estreitas a grade possui rolagem local; a página não transborda.
- `src/overview-presentation.ts`: apresentação adicional sobre os elementos originais, mantendo IDs e instâncias de gráficos. Move Receita × COGS para abaixo da grade e acrescenta progresso frente à meta mensal. Todos os demais painéis seguem disponíveis.
- `src/domain/metric-presentation.ts`: classificação visual de resultado, meta, comparação e progresso, com quatro testes específicos. Não calcula receitas, despesas nem projeções.

## Preservação

O script original de carga/cálculo permanece integral, byte a byte, dentro do modelo. O adaptador expõe um snapshot somente leitura dos totais e metas já calculados, sem mudar consultas, arredondamento, imposto, fallback, histórico ou fontes. O overview continua bloqueado para escrita. A DRE privada e os cenários mantêm seus fluxos e regras.

Resultado positivo/negativo, melhora/piora e meta são dimensões distintas. Aumento de despesa isolado fica neutro e informa custo maior/menor; CAC/CPV interpretam menor custo como melhora. Metas de receita/lucro e orçamento identificam expressamente seu escopo mensal. Margem mostra a referência original de 22%, sem inventar uma meta cadastrada. Metas zero são valores válidos; progresso percentual exige denominador positivo. Períodos que cruzam meses não são comparados a uma única meta mensal no novo painel. O resultado do painel informa que antecede despesas complementares da DRE.

Fallback de exemplo recebe aviso visível e identificação no cabeçalho. Cards sem dados ou sem meta não ganham números ou status artificiais. Temas claro/escuro, preferência de movimento reduzido e tabelas acessíveis de gráficos continuam funcionando.

## Validação

- Build TypeScript/Vite e release local, que reconstrói exclusivamente `console/` e verifica o hash da referência original.
- 144 testes aprovados: 132 de domínio/leitura (incluindo quatro novos de apresentação semântica) e 12 de equivalência/preservação.
- Navegador Chrome: 15 cards em cinco colunas, dados reais, cores negativas, filtro de período, exportação CSV, regras e Escape, rail móvel, Financeiro, Simulador, tema escuro; viewport de 390px sem overflow global, com rolagem local. Observação da view estável sem mutações contínuas.
- Evidências e scripts de revisão em `../../output/redesign-implementado-2026-10-29/` no workspace principal.

## Estado de publicação

Fonte e saída `console/` preparadas localmente. Não foi realizado push nem alterado o site público. Prévia local pelo Vite. Para publicar, seguir o fluxo existente do repositório, incluindo fonte e saída gerada; não modificar o dashboard HTML original nem o banco para essa troca de apresentação.

Rollback: reverter alterações de apresentação e regenerar `console/`. Não reverter tarefas, metas, despesas ou outros registros do operador.

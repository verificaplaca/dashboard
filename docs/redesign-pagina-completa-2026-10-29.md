# Visão Geral — acabamento da página completa

Continuação do design aprovado: fundo claro, superfícies brancas, controles em cápsula, tipografia leve e verde/coral com significado explícito. Os quinze indicadores permanecem em cinco colunas e três linhas, com 172px de altura e tags lado a lado.

## Composição

- Produtos em dois painéis iguais, consulta principal primeiro. Métricas agrupadas e gráficos alinhados na base. Ausência de CAC recebe uma nota explícita, sem preencher valores fictícios.
- Monitoramento de Ads e bureau em dois painéis com métricas de apoio em 2×2; histórico do saldo separado dos números principais; composição por fornecedor em blocos próprios.
- Estornos com leitura compacta e quantidade zero sem alarme vermelho. Valores continuam informativos, sem segunda dedução.
- Eficiência com barras discretas, referência e estado textual preservados. Upsell abaixo da meta em coral.
- Funil em três etapas horizontais no desktop e verticais em telas estreitas, mantendo quantidades, proporções e taxas originais.
- Resultado bruto/líquido pareados, com totais destacados, barras verdes/corais e média tracejada; CAC e upsell pareados abaixo, com metas originais.
- Médias da semana e distribuição por campanha em painéis próprios. Nota de atribuição explica as estimativas fora do teste.
- Todos os gráficos têm consulta aos valores em tabela; incluindo os dois gráficos de produtos.

## Integridade e validação

O script original de cálculo permanece congelado. As mudanças ocorrem na camada de apresentação, sem alterar séries, categorias, fórmulas, impostos, metas ou callbacks de formatação. O teto do gráfico de upsell é ampliado quando necessário para não esconder observações acima de 40%.

149 testes aprovados (137 de domínio/apresentação e 12 de paridade original). Build e release local reconstruídos; `git diff --check` sem erros. Verificação no navegador a 1448×1086: quinze cards, três linhas, altura uniforme de 172px e diferença subpixel de largura inferior a 0,01px. Em 390×844: três linhas mantidas, rolagem local do quadro financeiro e nenhum vazamento horizontal dos demais painéis. Tabela de valores do resultado bruto conferida, inclusive resultado diário negativo.

Evidências locais em `../../output/redesign-pagina-completa-2026-10-29/`. Não houve publicação remota.

## Refinamento de entendimento e interação
- Leitura do período calculada a partir dos totais originais; distingue resultado positivo, cumprimento da meta e melhoria do CAC.
- Detalhes de CAC, margem e upsell com fórmula, componentes, comparação e distância da meta. Margem distingue referência saudável de 22% da referência visual de 30%.
- Comparação opcional por janela anterior de mesma duração de calendário; lacunas permanecem ausentes, sem inventar zeros. Tooltip identifica data equivalente.
- Navegação sticky acompanha a seção ativa, respeita movimento reduzido e remove listeners ao desmontar.
- Produtos com metodologia recolhível; autonomia de Ads com contexto e acesso a campanhas.
- Transições curtas, foco devolvido ao acionador no shadow DOM, destaque do dia durante investigação e carregamento com espaço reservado.
- Os 15 cards existentes não foram redesenhados nesta etapa. Validação no navegador: 15 cards, altura uniforme de 172 px.
- Validação: 149 testes de domínio e 12 de paridade aprovados; build/release local concluídos; console sem erros na inspeção. Release preparado em console/, sem publicação remota.
- Evidências em output/redesign-pagina-completa-2026-10-29: cac-explicado.jpg e performance-comparativa.jpg (pasta no diretório pai do repositório).

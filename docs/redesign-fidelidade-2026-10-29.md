# Correção de fidelidade à referência

O usuário apontou que a primeira implementação não reproduzia a imagem aprovada e pediu cards idênticos e tags sem empilhamento.

## Correções implementadas

- Hero com título, descrição, período, atualização e exportação numa composição única. Barras adicionais removidas da leitura principal; filtros de campanha, atualização detalhada, atalhos de seção e regras permanecem no menu de três pontos.
- Sidebar reduzida aos destinos reais principais. Outros destinos continuam disponíveis em Todos os módulos/busca; marca oficial preservada, sem wordmark adicional.
- Quinze cards em 5×3, com altura fixa de 172px e colunas iguais. Slots explícitos para título, valor, contexto, estados, detalhe e minigráfico, sem altura variável por quantidade de texto.
- Tags de 16px de altura, fonte compacta, sem quebra de linha. Uma tag principal por indicador; CAC pode exibir comparação e meta lado a lado. Sinais de resultado positivo/negativo também estão no valor e na descrição acessível, sem uma terceira tag.
- Projeções, CAC Google Ads, fórmulas, comparativos e demais informações preservados nas informações de cada indicador. CSV continua exportando o comparativo original, mesmo quando ele não é a tag principal.
- Minigráficos em barras com as mesmas observações originais. Não foram usadas curvas ascendentes fictícias do mockup. Gráfico Receita/COGS com apresentação mais limpa; séries, projeções e valores originais preservados.
- Painel Meta de resultado aproximado da composição aprovada, mantendo período, meta mensal e limitações da DRE explícitos.

## Verificação

144 testes aprovados e release local reconstruído. Validação no navegador aberto, no viewport de referência 1448×1086: 15 cards; todos com 172px de altura; diferença de largura inferior a 0,01px (arredondamento subpixel do navegador); todos os slots internos alinhados; máximo de duas tags, sempre na mesma linha. Regras e exportação validadas após a reorganização.

Evidências em `../../output/redesign-implementado-2026-10-29/fidelidade-desktop.jpg` e `fidelidade-geometria.json`. O viewport usado na comparação foi restaurado após a verificação. Fonte e saída gerada preparadas localmente; sem publicação remota.

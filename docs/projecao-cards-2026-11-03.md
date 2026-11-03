# Projeção nos cards · 03/11/2026

A Visão Geral exibe a legenda `Proj.: … · Meta: …` em Receita Bruta, Custo Total COGS (com orçamento), Lucro Bruto e Lucro Líquido. A legenda vem diretamente de `ksub_*`, calculada pelo motor preservado. Apenas projeções disponibilizadas pelo original, em um único mês, são exibidas; os demais períodos mantêm a legenda de meta existente.

O rodapé ocupa uma linha própria abaixo do histórico. Cards desktop passam de 172 para 166px; miniaturas de 26 para 22px, com espessura máxima de barra de 6px para períodos curtos. No celular, o rodapé pode quebrar linha.

Validação: 226 testes de domínio e 12 de paridade aprovados, build/release local. Navegador em modo de demonstração da interface: 15 cards, legendas iguais às originais, miniaturas de 22px, sem erros de execução; viewport 390px sem transbordamento horizontal. Capturas em `output/projecao-cards-2026-11-03/`. Publicação autorizada em 03/11/2026 pelo fluxo automático da branch main para o projeto Vercel verifica-placa-ops. Base atualizada para 3bccc11 antes do build, preservando o login e a persistência de sessão recentes.

## Revisão: barras agrupadas e meta diária

As barras agora ficam à direita em uma faixa de até 5px por dia (4px de barra e aproximadamente 1px de espaço), limitada à largura disponível do card. Receita, COGS, lucro bruto e lucro líquido comparam a média realizada por dia com a parcela diária das metas mensais dos dias selecionados. Filtros entre meses respeitam a duração e a meta de cada mês; custos usam comparação inversa. A legenda da projeção continua mensal. O tooltip explica os valores e a regra. O lucro bruto passa a exibir o status diário, com a variação anterior preservada no tooltip; a cor do histórico do lucro líquido acompanha o ritmo diário.

Validação da revisão: 228 testes de domínio e 12 de paridade aprovados; release Pages e build Vercel aprovados. Navegador: 15 cards alinhados à direita em desktop e celular, projeções preservadas e nenhum erro de execução ou transbordamento horizontal. Evidências em `output/cards-meta-diaria-2026-11-03/`.

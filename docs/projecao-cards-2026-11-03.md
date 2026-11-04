# Projeção nos cards · 03/11/2026

A Visão Geral exibe a legenda `Proj.: … · Meta: …` em Receita Bruta, Custo Total COGS (com orçamento), Lucro Bruto e Lucro Líquido. A legenda vem diretamente de `ksub_*`, calculada pelo motor preservado. Apenas projeções disponibilizadas pelo original, em um único mês, são exibidas; os demais períodos mantêm a legenda de meta existente.

O rodapé ocupa uma linha própria abaixo do histórico. Cards desktop passam de 172 para 166px; miniaturas de 26 para 22px, com espessura máxima de barra de 6px para períodos curtos. No celular, o rodapé pode quebrar linha.

Validação: 226 testes de domínio e 12 de paridade aprovados, build/release local. Navegador em modo de demonstração da interface: 15 cards, legendas iguais às originais, miniaturas de 22px, sem erros de execução; viewport 390px sem transbordamento horizontal. Capturas em `output/projecao-cards-2026-11-03/`. Publicação autorizada em 03/11/2026 pelo fluxo automático da branch main para o projeto Vercel verifica-placa-ops. Base atualizada para 3bccc11 antes do build, preservando o login e a persistência de sessão recentes.

## Revisão: barras agrupadas e meta diária

As barras agora ficam à direita em uma faixa de até 5px por dia (4px de barra e aproximadamente 1px de espaço), limitada à largura disponível do card. Receita, COGS, lucro bruto e lucro líquido comparam a média realizada por dia com a parcela diária das metas mensais dos dias selecionados. Filtros entre meses respeitam a duração e a meta de cada mês; custos usam comparação inversa. A legenda da projeção continua mensal. O tooltip explica os valores e a regra. O lucro bruto passa a exibir o status diário, com a variação anterior preservada no tooltip; a cor do histórico do lucro líquido acompanha o ritmo diário.

Validação da revisão: 228 testes de domínio e 12 de paridade aprovados; release Pages e build Vercel aprovados. Navegador: 15 cards alinhados à direita em desktop e celular, projeções preservadas e nenhum erro de execução ou transbordamento horizontal. Evidências em `output/cards-meta-diaria-2026-11-03/`.

## Hover sobre a meta diária

O próprio indicador Dentro/Fora da meta diária agora tem tooltip nativo com duas linhas: Meta diária e Média realizada no período, em reais. Os valores usam o mesmo resultado do cálculo do status e se atualizam com o filtro; quando a referência diária não está disponível, o tooltip é removido. A descrição também é exposta na identificação acessível do indicador.

## Diferença no hover da meta diária

O tooltip também informa Acima/Abaixo da meta (ou Na meta), com a diferença absoluta entre a média diária realizada e a meta diária em reais e em percentual da meta. COGS mantém o sentido literal da diferença: gastar mais que o orçamento diário aparece como acima da meta e com status negativo. Meta zero não gera percentual infinito; informa que o percentual não é calculável. Validação: 229 testes de domínio e 12 de paridade, incluindo distância positiva/negativa, empate, custo, meta zero e meses com metas distintas.

## Remoção dos subtítulos dos cards

Retirados os textos explicativos curtos abaixo dos valores na Visão Geral e no Simulador (ex.: Ads + bureau). As explicações continuam nos tooltips/diálogos existentes. A grade foi compactada para eliminar a linha vazia: cards desktop da Visão Geral com 144px e do Simulador com 149px. Valores, projeções, metas e comparativos preservados.

Validação: 229 testes de domínio e 12 de paridade aprovados; builds Pages/Vercel concluídos. Navegador conferiu 15 cards sem subtítulos em cada tela, os 15 controles explicativos e a abertura da explicação do lucro bruto no Simulador. Visão Geral e Simulador sem transbordamento global em 390px após estabilização do layout. Evidências locais em `output/cards-sem-subtitulo-2026-11-03/`.

## Barras do Simulador no padrão da Visão Geral

Minigráficos do Simulador passam a ocupar uma faixa proporcional à quantidade de dias, alinhada à direita, limitada à largura disponível. Barra de 4px, intervalo aproximado de 1px e altura de 22px, seguindo a Visão Geral. Mantidos histórico, sinais negativos, lacunas e cores existentes.

Validação: 229 testes de domínio e 12 de paridade aprovados; builds Pages/Vercel concluídos. Navegador verificou alinhamento dos 15 cards e períodos de 14 dias (70px) e 7 dias (35px). Evidência em `output/simulador-barras-2026-11-03/`.

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

## Altura e espaçamento do Simulador iguais à Visão Geral

Cards do Simulador ajustados para 144px no desktop, padding de 14px × 18px e linhas internas de 18/34/18/22/16px, mantendo intervalo de 16px entre cards. Minigráfico na quarta linha e indicação de editável/calculado no rodapé, seguindo as posições da Visão Geral. Em telas pequenas, altura mínima de 173px e espaçamentos de 12/10/8px conforme o breakpoint, com crescimento conforme o conteúdo.

Validação: 229 testes de domínio e 12 de paridade aprovados; builds Pages/Vercel concluídos. Medidas conferidas no navegador no desktop e em 390px; mesma altura mínima, padding e intervalos da Visão Geral. Evidências em `output/simulador-espacamento-2026-11-03/`.

## Cores das barras do Simulador

Aplicados os tons exatos da Visão Geral: verde #a0f644, cinza #c6c9d1 e rosa #ff8290. COGS, ROAS, Ads, Bureau e Transações usam cinza; os demais indicadores usam verde. Perdas e piora de CAC, upsell ou ticket em relação ao controle do cenário usam rosa. Dias históricos negativos preservam sinalização rosa. O status parcial do lucro continua indicado no texto, sem introduzir uma quarta cor nas barras.

Validação: 229 testes de domínio e 12 de paridade aprovados; builds Pages/Vercel concluídos. Cores computadas dos 15 cards conferidas no navegador; CAC simulado em R$20 confirmou o rosa em CAC e lucro negativo, e cenário foi restaurado ao final. Evidência em `output/simulador-cores-2026-11-03/`.

## Cores dos valores do Simulador

Lucro bruto e líquido usam a mesma classificação de resultado da Visão Geral: positivo verde #07872d, negativo vermelho e zero neutro. Margem líquida segue verde para valor não negativo e vermelho para negativo; indisponibilidade permanece neutra. Demais valores usam a cor padrão da Visão Geral por tema (#111827 no claro e #edf2fa no escuro). Negativos usam #d92a3c no claro e #ff9ca8 no escuro. Removidas regras anteriores que pintavam apenas o lucro líquido ou substituíam a cor do resultado parcial; o aviso textual de resultado parcial continua disponível.

Validação: 229 testes de domínio e 12 de paridade aprovados; builds Pages/Vercel concluídos. Cores computadas dos 15 valores verificadas. Cenário temporário com CAC R$20 confirmou lucro negativo e tons corretos nos temas claro e escuro; cenário e tema restaurados ao final. Evidências em `output/simulador-valores-cores-2026-11-03/`.

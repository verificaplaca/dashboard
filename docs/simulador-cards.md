# Cards financeiros editáveis · 29/10/2026

A seção Resultados Financeiros do simulador reproduz os 15 cards e sua ordem na Visão Geral. Os números representam o cenário no horizonte escolhido (7/30/95 dias); não são os totais históricos da janela de calibração. Os minigráficos mostram o histórico real diário da base 7/14/30 dias, com lacunas e perdas preservadas. Os deltas comparam o cenário e manter a verba de referência no mesmo horizonte.

## O que editar

| Card | Regra da alteração |
|---|---|
| Receita Bruta | Escolher aumentar ticket mantendo pedidos, ou mudar pedidos mantendo ticket. |
| Custo Total COGS | Mantém Ads; redefine o custo unitário de bureau. Não aceita COGS menor que Ads. |
| Gasto Google Ads | Divide o total pelos dias e atualiza a verba diária, em centavos, até R$ 10 mil/dia. |
| Gasto Bureau | Redefine custo por pedido; o total acompanha o volume. |
| CAC | Define pedidos por real investido (1/CAC), para o período inteiro. |
| Transações | Define pedidos por real investido à verba atual; admite zero e preserva prejuízo. |
| Taxa de Conversão | Mantém custo por checkout histórico e recalcula pedidos; requer fonte completa. |
| Taxa de Upsell | Aplica somente a diferença de taxa ao ticket e bureau; exige receita/custo incrementais. |
| Ticket Médio | Muda receita mantendo volume; não estima causalmente reação a preço. |

Lucro Bruto, Lucro Líquido, ROAS, CPV e Margem são resultados calculados. Os cadeados abrem explicações; não permitem números independentes que deixariam a conta inconsistente. EBITDA mantém o acumulado histórico das mesmas fontes desde 01/01/2026 até ontem e não muda por hipótese. É a métrica rotulada assim no painel, não um EBITDA contábil completo.

## Contas e hipóteses

COGS = Ads + bureau. Lucro bruto = receita − COGS. ROAS = receita / COGS (definição deste painel, não receita / Ads). CPV = COGS / pedidos; CAC = Ads / pedidos; margem = lucro bruto / receita.

Lucro líquido do cenário = receita − Ads − bureau − imposto diário − taxas e despesas informadas na DRE. Despesas desconhecidas continuam pendentes e o destaque permanece parcial. A Visão Geral original mantém sua fórmula sem complementos privados; não foi modificada.

O controle sempre usa as relações históricas, sem as edições hipotéticas. Sem eficiência manual, as três respostas à verba continuam hipóteses de elasticidade 0,6/0,8/1, não regressões causais aprendidas da base. Adaptação conserva a hipótese de 6% de perda por sete dias após mudar a verba.

Informar CAC, transações, conversão ou receita via volume substitui o mesmo driver de aquisição; não empilha constraints incompatíveis. A eficiência passa a ser pedidos por real constante, com uma única curva. O CAC informado já é a eficiência final no horizonte: elasticidade e adaptação não são aplicadas novamente. Ao editar CAC/pedidos, preserva-se a conversão observada e calculam-se checkouts necessários, pressupondo mudança no custo do tráfego. Ao editar a conversão, mantém-se o custo por checkout e muda-se a eficiência. Razões sem denominador ou superiores a 100% ficam indisponíveis, sem criar uma taxa artificial. Contagens por data não constituem uma coorte de clientes.

Upsell usa `ticket_base + (taxa_nova − taxa_base) × receita_incremental`, com as taxas em fração; bureau segue a mesma diferença com custo incremental. A receita adicional por pedido com upsell pode ser sugerida por `upsell_daily.upsell_revenue / upsell_orders`. O custo incremental não está separado na fonte: o editor exige valor explícito, incluindo zero apenas quando confirmado. O valor atual do ticket já contém o mix histórico; somar todo o adicional novamente contaria receita duas vezes. Alterar preço/custo após um mix preserva a outra unidade e a hipótese de mix. Valores incompatíveis com as unidades médias são rejeitados.

O histórico calibra ticket, custo por pedido e razões observadas. Não determina elasticidade, impacto de preço ou custo incremental não registrado. Totais arredondados por dia podem diferir alguns centavos de um alvo digitado.

## Fontes e privacidade

GET nas fontes agregadas existentes. Core financeiro: mesmo `loadFinanceData`. Contexto: `upsell_daily` (pedidos pagos e receita), `checkouts_daily` no projeto original do site e acumulado financeiro de 2026. SQL de `upsell_daily` e `upsell_by_type` conferido somente em leitura: status paid/delivered/authorized, com mesma data de pagamento/São Paulo. Não se lê PII, pedidos individuais nem dados de usuário. Configuração do site é extraída da credencial anônima já pública no HTML original, com validação de papel anon e hostname fixo; acesso adicional do leitor ao projeto do site é limitado a checkouts_daily.

Fonte opcional indisponível não vira zero: o card afetado fica sem valor e não habilita edição que depende dele. Período/identidade atualizam o contexto; o histórico acumulado e as fontes reais não são alterados. Edições dos cards são rascunhos em memória, não gravam despesas, metas, snapshots avançados ou Ads. Trocar base mantém hipóteses explícitas; Restaurar cenário limpa edições e volta à verba de referência. Recarregar descarta este rascunho. As ferramentas avançadas e seus snapshots continuam separados e preservados.

## Verificação

140 testes (128 domínio/leitura e 12 paridade), TypeScript e release. Novos casos: conciliação das fórmulas, controle imutável, inversão da receita, CAC manual sem dupla penalidade, curva única, conversão/volume zero, upsell sem dupla contagem, preservação das unidades ao reeditar, custo desconhecido, limites, lacunas e fontes somente GET nos projetos corretos. O HTML original continua com o hash congelado; não há migração de banco nesta entrega.

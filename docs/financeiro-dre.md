# Financeiro e simulador simples · 29/10/2026

## Uso

O menu principal oferece **Financeiro** e **Simulador**. O Financeiro abre a DRE do mês, com vendas, impostos, consultas, anúncios e resultado. Troque o mês pelas setas ou pela lista; o mês atual termina ontem. “Baixar DRE” exporta os valores e pendências em CSV.

Taxas, equipe/pró-labore, sistemas, outras despesas e despesas financeiras/outros tributos são complementados pelo gestor autenticado. Campo vazio significa **não informado**; zero confirma ausência de custo. Enquanto houver pendências, o destaque é **Resultado parcial**, nunca lucro líquido completo. Não repetir custos de Ads, bureau ou impostos automáticos.

No Simulador, escolha a verba diária, a duração e se deseja considerar adaptação. Veja o resultado central, o ganho sobre manter e um gráfico com três hipóteses de crescimento. O pico destacado é matemático e limitado à faixa de R$ 0–10.000/dia, sem garantia de desempenho. A referência inicial é a verba da etapa registrada (R$ 2.790/dia no histórico inicial), editável nas considerações da conta. A base usa os últimos 7, 14 ou 30 dias encerrados. Ferramentas avançadas preservam cenários salvos, bases após 16/10 e a comparação até janeiro.

Mais verba aumenta lucro quando as vendas adicionais deixam contribuição maior que seu custo de aquisição. O CAC médio atual não garante o CAC das vendas adicionais. “Mesma eficiência” mantém resposta proporcional (elasticidade 1); a central usa 0,8; a prudente, 0,6. Não são probabilidades. Adaptação é uma hipótese de perda de 6% dos pedidos nos primeiros sete dias após mudar a verba; manter a referência não recebe nova penalidade.

## Fontes e reconhecimento

Leituras GET nas mesmas fontes da Visão Geral: `revenue_daily`, `bureau_daily`, `google_ads_campaign_daily`, `monthly_targets`; estornos informativos em `refunds_daily`. O imposto herda a última alíquota cadastrada; o fallback de 8% gera aviso e não completa a DRE. Falta de dia, duplicação ou custo inválido impedem os totais, em vez de criar zeros.

Conferência SQL em 29/10 confirmou: `revenue_daily` soma somente os status paid/delivered/authorized. Pedidos cancelados já são excluídos da receita. `refunds_daily` usa a data `updated_at` do cancelamento. Portanto, estornos são apresentados separadamente, **sem nova dedução**. Esta DRE gerencial depende da conciliação dessas fontes e não representa fluxo de caixa ou fechamento contábil oficial.

Despesas complementares entram integralmente no mês. No mês em curso, taxas correspondem às realizadas até ontem; demais despesas devem representar a competência do mês. Na projeção, taxas informadas viram proporção das vendas realizadas do mês e acompanham o faturamento projetado; demais despesas são rateadas pelos dias do mês e mantidas por dia. O resultado futuro continua parcial se faltarem valores. Despesas privadas de um mês não são inventadas para outros meses.

## Persistência privada

Migração aditiva `021_console_finance.sql`: `console_finance_months` e histórico versionado `console_finance_history`. Apenas o gestor ativo já cadastrado em `console_task_members` pode ler/salvar. Nenhum novo usuário ou permissão de metas foi criado. Anônimo não pode ler as tabelas nem executar as RPCs. Custos privados não entram no HTML, na configuração pública ou nos snapshots avançados. A DRE pública usa apenas as fontes agregadas que o sistema já expõe.

Cliente autenticado tem SELECT sob RLS e EXECUTE nas RPCs; não tem INSERT/UPDATE/DELETE direto. A RPC valida identidade, membership, categorias exatas e valores não negativos com até duas casas decimais; bloqueia membership durante a gravação; exige a versão observada e grava histórico na mesma transação. Conflito mantém o rascunho e exige carregar os valores atuais, sem sobrescrita automática. O mês e a identidade indexam as leituras para impedir exibir custos de outro mês/usuário durante a troca.

A migração foi aplicada individualmente e atomicamente no projeto analytics vinculado. Não executar `db push` em bloco sobre o legado. Testes SQL transacionais cobrem anonimato, gestor, outsider, revogação, validação, versões e histórico com ROLLBACK. Não foram inseridas despesas reais ou modificados orçamentos Ads. O salvamento visual pela conta do gestor exige seu login normal; não há sessões de impersonação ou senha em chat.

## Verificação e manutenção

129 testes automatizados (117 domínio/leitura e 12 paridade) cobrem a DRE, custos desconhecidos versus zero, estornos, meses, fontes somente leitura, margem incremental, adaptação, máximo limitado, permissões e conflitos. `npm --prefix console-app run release` valida TypeScript, build e preservação do HTML original, e atualiza exclusivamente `console/`. Fonte e saída devem ser commitadas juntas.

Rollback de interface: reverter o commit do frontend e regenerar `console/`. Manter tabelas e histórico privados; não apagar despesas posteriores do gestor. A interface anterior do simulador permanece em `AdvancedSimulation.tsx`, carregada sob demanda, com os mesmos identificadores locais de snapshots e etapas.

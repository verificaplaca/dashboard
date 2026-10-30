# Início · resumo da operação

A inicial de produção (`#/inicio`) tem composição própria em `OperationsHome.tsx` e CSS isolado em `operations-home.css`. Seus quatro indicadores usam exatamente os totais do motor preservado da Visão Geral: Receita Bruta (`rev`), Lucro Líquido (`netPftOpt`), Transações (`conv`) e CAC (`cac`). O snapshot também fornece investimento, metas e receita diária.

## Período e cálculo

O mesmo filtro salvo em `vp_filter` rege os dois módulos. O mês atual inclui hoje; uma seleção de outro intervalo em qualquer um dos módulos é respeitada no outro. Na inicial, o botão de período abre os mesmos atalhos da Visão Geral e um intervalo personalizado com datas inicial/final. Aplicar atualiza o resumo sem sair da página. O cabeçalho mostra o período e indica quando inclui hoje. Os dados do dia seguem sujeitos a atualização das fontes.

`OriginalOverview` executa o motor original em DOM isolado, fora da tela, inerte e oculto da árvore de acessibilidade. A ponte `readHomeSnapshot` lê `totals()`, `activeTargets()` e os dias filtrados; não duplica nem altera as fórmulas originais. O usuário continua vendo apenas a composição própria da inicial. O lucro segue a Visão Geral, após imposto, Ads e bureau; despesas complementares da DRE pertencem ao Financeiro.

Os até sete dias mais recentes do filtro formam o gráfico. A meta é mensal: períodos com mais de um mês não exibem percentual de realização contra uma única meta. Falha ou dados de exemplo não são apresentados como indicadores reais. Atualizar resumo relê o motor, tracking e tarefas, com loader global e saída suave.

## Outros blocos

Tarefas obedecem à conta autenticada. Tracking consulta os últimos 50 envios em sete dias de calendário e explicita essa amostra. Falha tem prioridade sobre pendência e status desconhecido fica a conferir; `skipped` não é confirmação de entrega. Os atalhos mantêm suas permissões; Usuários aparece apenas para administradores.

Metas podem ser consultadas por gerentes. Somente administradores ativos podem alterá-las, inclusive o imposto. A interface, o salvamento e as políticas do banco verificam essa regra (migração 023).

## Validação

184 testes da aplicação aprovados: 172 de domínio e 12 de paridade do motor original. O teste do snapshot executa a função original com ontem e hoje e compara cada total. Build TypeScript/Vite aprovado. Testes SQL locais e remotos cobrem escrita de administrador, leitura de gerente e bloqueio de mutações por gerente, metadata forjada, administrador inativo e anônimo, com ROLLBACK. Conferência remota: zero contas e metas de teste persistidas; duas memberships reais preservadas.

Intervalos personalizados sem registros preservam a seleção ao voltar ou recarregar e mostram valores zerados. A ponte protege o caso vazio que o renderizador original não previa, sem mudar fórmulas financeiras.

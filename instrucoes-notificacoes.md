# Instruções de deploy — melhorias de notificações Telegram

## 1. Rodar a migration

No SQL Editor do projeto **ftmgmfdqdqxboiktxcoj** (Dashboard), rodar:

```
supabase/migrations/012_notifications.sql
```

Isso adiciona:
- `ads_balance_history.alert_sent` (boolean, default false)
- tabela `health_alerts` (dedup de alertas do watchdog), RLS habilitado sem policies (só service_role acessa)

## 2. Deploy das functions

```bash
supabase functions deploy check-ads-balance --project-ref ftmgmfdqdqxboiktxcoj
supabase functions deploy resumo-diario --project-ref ftmgmfdqdqxboiktxcoj
supabase functions deploy check-tracking-health --project-ref ftmgmfdqdqxboiktxcoj
```

## 3. cron-job.org

- **Criar job novo**: `resumo-diario`
  - URL: `POST https://ftmgmfdqdqxboiktxcoj.supabase.co/functions/v1/resumo-diario`
  - Header: `Authorization: Bearer <SERVICE_ROLE_KEY>`
  - Schedule: `0 11 * * *` (08:00 BRT)

- **Conferir que continuam ativos** (nenhuma mudança de schedule neles):
  - `check-ads-balance` — 1x/hora
  - `check-tracking-health` — 1x/hora

## 4. Aposentar check-bureau-spend

`check-bureau-spend` está morto desde 19/05 — o `resumo-diario` absorve o papel dele (custo de bureau agora sai no resumo diário). Se existir job dele no cron-job.org, pode excluir.

## 5. Recomendação

Ativar notificação por e-mail de falha de job nas configurações da conta do cron-job.org (para saber se algum dos crons acima parar de rodar).

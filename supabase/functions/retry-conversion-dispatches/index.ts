/**
 * retry-conversion-dispatches
 *
 * Job de retry pro dispatch de conversões (GA4 MP / Google Ads / Meta CAPI).
 * Varre conversion_dispatches por linhas com algum canal != 'success' e
 * attempts < 5, e re-tenta SÓ os canais que falharam (não reenvia os que já
 * tiveram sucesso — evita duplicar Purchase em GA4/Meta).
 *
 * Não precisa rechamar o RPC do site: email_sha256/phone_sha256 (P2.3 — hash
 * at rest, sem PII em claro) já estão salvos em conversion_dispatches desde
 * o primeiro dispatch.
 *
 * Deploy: supabase functions deploy retry-conversion-dispatches --project-ref ftmgmfdqdqxboiktxcoj
 * Cron:   cron-job.org → POST .../functions/v1/retry-conversion-dispatches, a cada 15min
 *         Header: Authorization: Bearer <SERVICE_ROLE_KEY>
 */

import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'
import { dispatchAll } from '../_shared/conversion-dispatch.ts'

const MAX_ATTEMPTS = 5

Deno.serve(async (req) => {
  const supabase = createClient(
    Deno.env.get('SUPABASE_URL')!,
    Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
  )

  const { data: pending, error } = await supabase
    .from('conversion_dispatches')
    .select('*')
    .lt('attempts', MAX_ATTEMPTS)
    // Só failed/pending — 'skipped' (canal não configurado, ex: Meta sem pixel/token)
    // não é falha e não deve ser re-tentado (evita queimar attempts e evita
    // Purchase retroativo quando o canal for configurado no futuro).
    .or('ga4_status.in.(failed,pending),ads_status.in.(failed,pending),meta_status.in.(failed,pending)')
    .limit(50)

  if (error) return json({ ok: false, error: error.message }, 500)
  if (!pending?.length) return json({ ok: true, retried: 0 })

  let retried = 0
  for (const row of pending) {
    const result = await dispatchAll({
      checkout_id: row.checkout_id,
      order_nsu:   row.order_nsu,
      valor:       row.value,
      // P2.3 — conversion_dispatches não guarda mais email/phone cru, só o
      // hash (email_sha256/phone_sha256); os dispatchers usam o hash pronto.
      email_sha256: row.email_sha256,
      phone_sha256: row.phone_sha256,
      paid_at:     row.paid_at,
      // P1.4 — atribuição persistida em conversion_dispatches no primeiro
      // dispatch (migration 010); reaproveitada aqui sem precisar rechamar
      // a RPC do site.
      gclid:             row.gclid,
      gbraid:            row.gbraid,
      wbraid:            row.wbraid,
      fbp:               row.fbp,
      fbc:               row.fbc,
      ga_client_id:      row.ga_client_id,
      event_source_url:  row.event_source_url,
      client_user_agent: row.client_user_agent,
    })

    // Só sobrescreve o status de canais que ainda não tinham tido sucesso —
    // se ga4 já era 'success', mantém (não reenvia, evita duplicar).
    const update: Record<string, unknown> = {
      attempts: row.attempts + 1,
      updated_at: new Date().toISOString(),
    }
    if (row.ga4_status !== 'success')  { update.ga4_status  = result.ga4.status;  update.ga4_error  = result.ga4.error  ?? null }
    if (row.ads_status !== 'success')  { update.ads_status  = result.ads.status;  update.ads_error  = result.ads.error  ?? null }
    if (row.meta_status !== 'success') { update.meta_status = result.meta.status; update.meta_error = result.meta.error ?? null }

    await supabase.from('conversion_dispatches').update(update).eq('checkout_id', row.checkout_id)
    retried++
  }

  return json({ ok: true, retried })
})

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body, null, 2), { status, headers: { 'Content-Type': 'application/json' } })
}

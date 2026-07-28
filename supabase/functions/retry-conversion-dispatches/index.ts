/**
 * retry-conversion-dispatches
 *
 * Job de retry pro dispatch de conversões (GA4 MP / Google Ads / Meta CAPI).
 * Varre conversion_dispatches por linhas com algum canal != 'success' e
 * attempts < 5, e re-tenta SÓ os canais que falharam (não reenvia os que já
 * tiveram sucesso — evita duplicar Purchase em GA4/Meta).
 *
 * Não precisa rechamar o RPC do site: email_sha256 e phone_sha256_e164/_digits
 * (P2.3 — hash at rest, sem PII em claro) já estão salvos em conversion_dispatches
 * desde o primeiro dispatch. Consequência: linha gravada pelo CATCH do
 * dispatch-conversions (RPC do site falhou → sem hash nenhum, event_id='unknown')
 * NÃO é recuperável aqui — quem reprocessa essas é o reconcile.
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
    // Dispara SÓ os canais que ainda não deram success. Antes, dispatchAll
    // reenviava os três sempre e o código abaixo só evitava sobrescrever o
    // status — o que na prática mandava purchase repetido pro GA4 MP a cada
    // tentativa em qualquer linha com falha parcial de outro canal.
    const result = await dispatchAll({
      checkout_id: row.checkout_id,
      order_nsu:   row.order_nsu,
      valor:       row.value,
      // P2.3 — conversion_dispatches não guarda mais email/phone cru, só o hash.
      // Telefone em dois formatos (migration 015): o Google precisa do E.164 e o
      // Meta dos dígitos. `phone_sha256` é a coluna legada (dígitos), passada
      // adiante só pro Meta aproveitar linhas anteriores à migration.
      email_sha256:        row.email_sha256,
      phone_sha256_e164:   row.phone_sha256_e164,
      phone_sha256_digits: row.phone_sha256_digits,
      phone_sha256:        row.phone_sha256,
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
    }, {
      ga4:  row.ga4_status  !== 'success',
      ads:  row.ads_status  !== 'success',
      meta: row.meta_status !== 'success',
    })

    // Canal já 'success' não foi disparado acima e também não tem status
    // sobrescrito aqui (dispatchAll devolve 'skipped' pra ele).
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

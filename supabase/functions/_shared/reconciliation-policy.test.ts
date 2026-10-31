import { dispatchReconciliationCandidate, RECONCILIATION_GRACE_MS, type ExistingDispatch } from './reconciliation-policy.ts'
import type { CheckoutConversionData, DispatchChannels } from './conversion-dispatch.ts'
const assert = (value: unknown, message = 'Assertion failed') => { if (!value) throw new Error(message) }
const now = Date.parse('2026-10-31T12:00:00Z')
const data: CheckoutConversionData = { checkout_id: 'fixture-checkout', order_nsu: 'addon_fixture', valor: 64.90, paid_at: '2026-10-31T11:00:00Z' }
const healthy: ExistingDispatch = { event_id: 'evt_addon_fixture', email_sha256: 'fixture-hash', ga4_status: 'success', ads_status: 'success', meta_status: 'success' }
function fixture(existing: ExistingDispatch | null = null, error: unknown = null) {
  let reads = 0
  const calls: { data: CheckoutConversionData; channels: DispatchChannels }[] = []
  return {
    read: async () => { reads++; return { data: existing, error } },
    send: async (value: CheckoutConversionData, channels: DispatchChannels) => {
      calls.push({ data: value, channels })
      return { eventId: `evt_${value.order_nsu}`, ga4: { status: channels.ga4 ? 'success' as const : 'skipped' as const }, ads: { status: channels.ads ? 'success' as const : 'skipped' as const }, meta: { status: channels.meta ? 'success' as const : 'skipped' as const } }
    }, calls, reads: () => reads,
  }
}
Deno.test('payment racing the webhook is deferred without any platform request or database status write', async () => {
  // Observed overlap: webhook and recovery were active within seven seconds
  // of the same payment. This must no longer produce a second upload.
  const f = fixture()
  assert(await dispatchReconciliationCandidate({ ...data, paid_at: new Date(now - 7_000).toISOString() }, f.read, f.send, now) === null)
  assert(f.calls.length === 0 && f.reads() === 0)
})
Deno.test('a webhook completed after the initial bulk scan is not replayed', async () => {
  const f = fixture(healthy)
  assert(await dispatchReconciliationCandidate(data, f.read, f.send, now) === null)
  assert(f.reads() === 1 && f.calls.length === 0)
})
Deno.test('genuinely missing older payments still recover with the same IDs, amounts and all channels', async () => {
  const f = fixture()
  const result = await dispatchReconciliationCandidate(data, f.read, f.send, now)
  assert(result?.eventId === 'evt_addon_fixture' && f.calls.length === 1 && f.calls[0].data === data)
  assert(Object.values(f.calls[0].channels).every(Boolean))
})
Deno.test('broken conversion metadata is repaired without resending channels already successful', async () => {
  const f = fixture({ ...healthy, event_id: 'unknown', ads_status: 'failed' })
  const result = await dispatchReconciliationCandidate(data, f.read, f.send, now)
  assert(f.calls.length === 1 && f.calls[0].channels.ga4 === false && f.calls[0].channels.meta === false && f.calls[0].channels.ads === true)
  assert(result?.ga4.status === 'success' && result?.meta.status === 'success' && result?.ads.status === 'success')
})
Deno.test('failed freshness reads never cause a speculative duplicate recovery', async () => {
  const f = fixture(null, new Error('read failed'))
  let rejected = false
  try { await dispatchReconciliationCandidate(data, f.read, f.send, now) } catch { rejected = true }
  assert(rejected && f.calls.length === 0)
})
Deno.test('the grace boundary does not remove older missing orders from recovery', async () => {
  const f = fixture()
  await dispatchReconciliationCandidate({ ...data, paid_at: new Date(now - RECONCILIATION_GRACE_MS).toISOString() }, f.read, f.send, now)
  assert(f.calls.length === 1)
})

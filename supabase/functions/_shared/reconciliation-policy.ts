import type { CheckoutConversionData, DispatchChannels, DispatchResult } from './conversion-dispatch.ts'

// Reconciliation is a recovery path, not a second sender for fresh payments.
// Give the webhook time to finish before attempting recovery. Deferred orders
// remain eligible for the existing 12-hour lookback on the next scheduled run.
export const RECONCILIATION_GRACE_MS = 10 * 60_000

export interface ExistingDispatch {
  event_id: string | null
  email_sha256?: string | null
  phone_sha256_e164?: string | null
  phone_sha256_digits?: string | null
  phone_sha256?: string | null
  ga4_status: string; ga4_error?: string | null
  ads_status: string; ads_error?: string | null
  meta_status: string; meta_error?: string | null
}

export function needsConversionRepair(row: Pick<ExistingDispatch, 'event_id' | 'email_sha256' | 'phone_sha256_e164' | 'phone_sha256_digits' | 'phone_sha256'>): boolean {
  return row.event_id === 'unknown' || (!row.email_sha256 && !row.phone_sha256_e164 && !row.phone_sha256_digits && !row.phone_sha256)
}

export async function dispatchReconciliationCandidate(
  data: CheckoutConversionData,
  readCurrent: () => Promise<{ data: ExistingDispatch | null; error: unknown }>,
  send: (data: CheckoutConversionData, channels: DispatchChannels) => Promise<DispatchResult & { eventId: string }>,
  now = Date.now(),
): Promise<(DispatchResult & { eventId: string }) | null> {
  const paidAt = data.paid_at ? Date.parse(data.paid_at) : NaN
  if (Number.isFinite(paidAt) && now - paidAt < RECONCILIATION_GRACE_MS) return null

  // The bulk scan may be stale by the time the site RPC completes. A webhook
  // that finished since that scan must not be dispatched again.
  const current = await readCurrent()
  if (current.error) throw new Error('Could not verify the current conversion dispatch; recovery deferred')
  if (current.data && !needsConversionRepair(current.data)) return null

  const channels = {
    ga4: current.data?.ga4_status !== 'success',
    ads: current.data?.ads_status !== 'success',
    meta: current.data?.meta_status !== 'success',
  }
  const result = await send(data, channels)
  // A repaired row may already have a successful channel. Keep its actual
  // status instead of persisting the dispatcher's “not requested” placeholder.
  for (const key of ['ga4', 'ads', 'meta'] as const) {
    if (!channels[key]) result[key] = { status: 'success', ...(current.data?.[`${key}_error`] ? { error: current.data[`${key}_error`]! } : {}) }
  }
  return result
}

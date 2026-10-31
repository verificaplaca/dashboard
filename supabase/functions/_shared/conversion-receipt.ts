/** Diagnostic only: never changes a delivery decision or retries an event. */
export async function observeMetaReceipt(
  response: Response,
  eventId: string,
  value: number,
  paidAt: string | null | undefined,
  log: (message: string) => void = console.info,
): Promise<void> {
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    const body = await Promise.race([
      response.clone().json(),
      new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new Error('receipt timeout')), 1500) }),
    ])
    const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(eventId))
    const eventHash = Array.from(new Uint8Array(digest), byte => byte.toString(16).padStart(2, '0')).join('')
    log(JSON.stringify({
      kind: 'meta_capi_receipt_v1',
      event_hash: eventHash,
      http_status: response.status,
      events_received: typeof body?.events_received === 'number' ? body.events_received : null,
      expected_events: 1,
      value,
      currency: 'BRL',
      paid_day_utc: paidAt && Number.isFinite(Date.parse(paidAt)) ? new Date(paidAt).toISOString().slice(0, 10) : null,
      error_code: typeof body?.error?.code === 'number' ? body.error.code : null,
      message_count: Array.isArray(body?.messages) ? body.messages.length : 0,
    }))
  } catch {
    // An unavailable/malformed acknowledgement must not turn a successful
    // production upload into a failure and trigger an unnecessary replay.
    try { log(JSON.stringify({ kind: 'meta_capi_receipt_v1', http_status: response.status, receipt_unavailable: true })) } catch { /* diagnostics never break dispatch */ }
  } finally {
    if (timer !== undefined) clearTimeout(timer)
  }
}

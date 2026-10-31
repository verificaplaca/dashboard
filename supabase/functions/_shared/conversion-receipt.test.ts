import { observeMetaReceipt } from './conversion-receipt.ts'
import { dispatchAll } from './conversion-dispatch.ts'

function assert(value: unknown, message = 'Assertion failed'): asserts value { if (!value) throw new Error(message) }

Deno.test('acknowledgements retain only bounded diagnostic fields, not payloads or platform messages', async () => {
  const logs: string[] = []
  const response = Response.json({ events_received: 1, messages: ['private detail'], fbtrace_id: 'private trace', error: { code: 100, message: 'private error' } })
  await observeMetaReceipt(response, 'evt_order_fixture', 14.99, '2026-10-31T13:00:00Z', message => logs.push(message))
  const record = JSON.parse(logs[0])
  assert(record.events_received === 1 && record.expected_events === 1 && record.value === 14.99 && record.currency === 'BRL')
  assert(/^[a-f0-9]{64}$/.test(record.event_hash) && record.paid_day_utc === '2026-10-31')
  assert(!logs[0].includes('evt_order_fixture') && !logs[0].includes('private'))
  assert((await response.json()).events_received === 1, 'observing must not consume the original response')
})

Deno.test('zero acknowledgements remain observable without inventing acceptance', async () => {
  let record: Record<string, unknown> = {}
  await observeMetaReceipt(Response.json({ events_received: 0 }), 'fixture', 9.90, null, message => { record = JSON.parse(message) })
  assert(record.events_received === 0 && record.paid_day_utc === null)
})

Deno.test('malformed JSON and broken logging never escape into the dispatcher', async () => {
  const messages: string[] = []
  await observeMetaReceipt(new Response('not json'), 'fixture', 12.90, undefined, message => messages.push(message))
  assert(JSON.parse(messages[0]).receipt_unavailable === true)
  await observeMetaReceipt(Response.json({ events_received: 1 }), 'fixture', 12.90, 'invalid', () => { throw new Error('logging unavailable') })
})

Deno.test('receipt observation preserves exact Meta request IDs, values and status; GA4 and Ads remain unrequested', async () => {
  const originalFetch = globalThis.fetch, originalLog = console.info
  const env = ['META_PIXEL_ID', 'META_ACCESS_TOKEN'].map(name => ({ name, value: Deno.env.get(name) }))
  const requests: { url: string; body: Record<string, unknown> }[] = [], logs: string[] = []
  try {
    Deno.env.set('META_PIXEL_ID', 'fixture-pixel'); Deno.env.set('META_ACCESS_TOKEN', 'fixture-token')
    console.info = message => logs.push(String(message))
    globalThis.fetch = ((input, init) => {
      requests.push({ url: String(input), body: JSON.parse(String(init?.body)) })
      return Promise.resolve(Response.json({ events_received: 1 }))
    }) as typeof fetch
    for (const [order, value] of [['order_fixture', 14.99], ['addon_fixture', 64.90]] as const) {
      const result = await dispatchAll({ checkout_id: 'fixture-checkout', order_nsu: order, valor: value, paid_at: '2026-10-31T13:00:00Z', ga_client_id: '123.456' }, { ga4: false, ads: false })
      assert(result.meta.status === 'success' && result.ga4.status === 'skipped' && result.ads.status === 'skipped')
      const event = (requests.at(-1)!.body.data as Record<string, unknown>[])[0]
      const data = event.custom_data as Record<string, unknown>
      assert(event.event_name === 'Purchase' && event.event_id === `evt_${order}` && event.event_time === 1793451600)
      assert(data.order_id === order && data.value === value && data.currency === 'BRL')
    }
    assert(requests.length === 2 && requests.every(request => request.url.startsWith('https://graph.facebook.com/v20.0/fixture-pixel/events?')))
    assert(logs.length === 2 && logs.every(message => JSON.parse(message).events_received === 1))
  } finally {
    globalThis.fetch = originalFetch; console.info = originalLog
    for (const { name, value } of env) { if (value === undefined) Deno.env.delete(name); else Deno.env.set(name, value) }
  }
})

Deno.test('all three real dispatch paths keep their canonical order and value contracts', async () => {
  const originalFetch = globalThis.fetch, originalLog = console.info
  const names = ['GA4_MEASUREMENT_ID', 'GA4_API_SECRET', 'META_PIXEL_ID', 'META_ACCESS_TOKEN', 'GADS_CUSTOMER_ID', 'GADS_DEVELOPER_TOKEN', 'GADS_REFRESH_TOKEN', 'GADS_CLIENT_ID', 'GADS_CLIENT_SECRET', 'GADS_PURCHASE_CONVERSION_ACTION_ID', 'GADS_CLICK_CONVERSION_ACTION_ID', 'GADS_EC_SEND_TIMESTAMPS', 'GADS_EC_SEND_CONVERSION_DATE_TIME']
  const original = names.map(name => ({ name, value: Deno.env.get(name) }))
  const requests: { url: string; body: Record<string, unknown> }[] = []
  try {
    for (const name of names) Deno.env.delete(name)
    for (const name of names.slice(0, 10)) Deno.env.set(name, `fixture-${name}`)
    console.info = () => {}
    globalThis.fetch = ((input, init) => {
      const url = String(input)
      if (url === 'https://oauth2.googleapis.com/token') return Promise.resolve(Response.json({ access_token: 'fixture-oauth' }))
      requests.push({ url, body: JSON.parse(String(init?.body)) })
      if (url.startsWith('https://www.google-analytics.com/')) return Promise.resolve(new Response(null, { status: 204 }))
      return Promise.resolve(Response.json(url.startsWith('https://graph.facebook.com/') ? { events_received: 1 } : { results: [{}] }))
    }) as typeof fetch
    const result = await dispatchAll({ checkout_id: 'fixture-checkout', order_nsu: 'addon_fixture', valor: 64.90, paid_at: '2026-10-31T13:00:00Z', ga_client_id: '123.456', email_sha256: 'fixture-email-hash' })
    assert(result.ga4.status === 'success' && result.ads.status === 'success' && result.meta.status === 'success')
    assert(requests.length === 3, 'one request per destination; observation must not resend')
    const ga = requests.find(r => r.url.startsWith('https://www.google-analytics.com/'))!.body
    const event = (ga.events as { params: Record<string, unknown> }[])[0].params
    assert(ga.client_id === '123.456' && ga.timestamp_micros === 1793451600000000)
    assert(event.transaction_id === 'addon_fixture' && event.event_id === 'evt_addon_fixture' && event.value === 64.90 && event.currency === 'BRL')
    const ads = requests.find(r => r.url.includes('uploadConversionAdjustments'))!.body
    const adjustment = (ads.conversionAdjustments as Record<string, unknown>[])[0]
    assert(adjustment.orderId === 'addon_fixture' && adjustment.adjustmentType === 'ENHANCEMENT')
    assert(!('restatementValue' in adjustment) && !('conversionValue' in adjustment), 'existing Ads values must not be changed by diagnostics')
  } finally {
    globalThis.fetch = originalFetch; console.info = originalLog
    for (const { name, value } of original) { if (value === undefined) Deno.env.delete(name); else Deno.env.set(name, value) }
  }
})

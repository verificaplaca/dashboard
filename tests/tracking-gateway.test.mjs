import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import test from 'node:test';
import assert from 'node:assert/strict';
const html = readFileSync(new URL('../tracking-gateway.html', import.meta.url), 'utf8');
const code = html.match(/<script>([\s\S]*?)<\/script>/)[1].replace(/loadAndRender\(\);\s*$/, '');
function harness() {
  const elements = new Map();
  const element = key => {
    if (!elements.has(key)) elements.set(key, { innerHTML: '', textContent: '', style: {}, classList: { add() {}, remove() {} }, addEventListener() {} });
    return elements.get(key);
  };
  const context = vm.createContext({ console, setTimeout: cb => cb(), clearTimeout() {}, Date, URLSearchParams, navigator: {}, requestAnimationFrame: cb => cb(), window: { innerWidth: 1200, addEventListener() {}, parent: { postMessage() {} } }, document: { getElementById: element, querySelector: element, querySelectorAll: () => [], addEventListener() {}, body: {}, documentElement: {} }, ResizeObserver: class { observe() {} } });
  vm.runInContext(code.replace(/\/\/ ─── BOOTSTRAP[\s\S]*/, ''), context);
  vm.runInContext(`
    var rendered = [];
    renderLastSync = () => rendered.push('sync');
    applyFilters = () => rendered.push('orders');
    renderHealth = () => rendered.push('health');
    renderChannels = () => rendered.push('channels');
    renderAttribution = () => rendered.push('attribution');
    renderChart = () => rendered.push('chart');
    renderCac = () => rendered.push('cac');
  `, context);
  return { context, element, run: sql => vm.runInContext(sql, context) };
}
test('revenue timeout preserves real tracking and clears stale CAC', async () => {
  const h = harness();
  h.element('#cacTable tbody').innerHTML = 'old data';
  h.run(`supaGet = async path => { if (path.includes('revenue_daily')) throw Object.assign(new Error('statement timeout'), {status:500,source:'revenue_daily'}); return path.includes('conversion_dispatches') ? [{checkout_id:'real'}] : []; };`);
  await h.run('loadAndRender()');
  assert.deepEqual(Array.from(h.run('rendered')), ['sync','orders','health','channels','attribution','chart']);
  assert.equal(h.run('STATE.orders[0].checkout_id'), 'real');
  assert.match(h.element('cacSummaryRow').innerHTML, /Comparativo CAC indisponível/);
  assert.equal(h.element('#cacTable tbody').innerHTML, '');
});
test('tracking outage clears previous records and does not render a false healthy zero', async () => {
  const h = harness();
  h.run(`STATE.orders = [{checkout_id:'old'}]; supaGet = async path => { if (path.includes('conversion_dispatches')) throw Object.assign(new Error('denied'), {status:403,source:'conversion_dispatches_public'}); return []; };`);
  await h.run('loadAndRender()');
  assert.equal(h.run('STATE.orders.length'), 0);
  assert.deepEqual(Array.from(h.run('rendered')), ['cac']);
  for (const id of ['healthRow','channelsRow','attrRow','chartSvgWrap']) assert.match(h.element(id).innerHTML, /Tracking indisponível/);
});
test('stale load cannot overwrite newer tracking or CAC', async () => {
  const h = harness();
  h.run(`var pending = []; supaGet = path => new Promise((resolve,reject) => pending.push({resolve,reject,path}));`);
  const old = h.run('loadAndRender()');
  const current = h.run('loadAndRender()');
  h.run(`pending[3].resolve([{checkout_id:'new'}]); pending[4].resolve([]); pending[5].resolve([]);`);
  await current;
  h.run(`pending[0].resolve([{checkout_id:'old'}]); pending[1].reject(new Error('old timeout')); pending[2].resolve([]);`);
  await old;
  assert.equal(h.run('STATE.orders[0].checkout_id'), 'new');
  assert.equal(h.element('cacSummaryRow').innerHTML, '');
});
test('authorization errors are not retried and diagnostics escape source content', async () => {
  const h = harness();
  h.context.fetch = async () => ({ ok:false, status:403, json:async () => ({code:'42501',message:'<script>denied</script>'}) });
  let calls = 0;
  const fetch = h.context.fetch;
  h.context.fetch = async (...args) => { calls++; return fetch(...args); };
  await assert.rejects(h.run("supaGet('/rest/v1/revenue_daily?select=date')"), /42501/);
  assert.equal(calls, 1);
  assert.match(h.run("sourceError('Unavailable', new Error('<script>denied</script>'))"), /&lt;script&gt;/);
});
test('tracking remains visible before the financial comparison resolves', async () => {
  const h = harness();
  h.run(`var delayed; supaGet = async path => path.includes('revenue_daily') ? new Promise(resolve => delayed=resolve) : path.includes('conversion_dispatches') ? [{checkout_id:'real'}] : [];`);
  const loading = h.run('loadAndRender()');
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(h.run('STATE.orders[0].checkout_id'), 'real');
  assert.ok(h.run("rendered.includes('health')"));
  assert.equal(h.run("rendered.includes('cac')"), false);
  h.run('delayed([])');
  await loading;
  assert.equal(h.run("rendered.includes('cac')"), true);
});

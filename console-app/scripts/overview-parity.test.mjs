import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import ts from 'typescript';

const original = fs.readFileSync(new URL('../reference/dashboard-original.html', import.meta.url), 'utf8');
const generated = fs.readFileSync(new URL('../public/overview/index.html', import.meta.url), 'utf8');
const inlineScripts = html => [...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)].filter(match => !/\bsrc\s*=/.test(match[1])).map(match => match[2]);
const calculationScript = html => inlineScripts(html).filter(script => script.includes('function totals()') && script.includes('async function loadSupabaseData()')).at(-1);
const upstream = calculationScript(original), embedded = calculationScript(generated);
const readModel = () => JSON.parse(fs.readFileSync(new URL('../public/overview/model.json', import.meta.url), 'utf8'));

// Extract declarations without executing initializers, initDashboard, DOM code or fetch.
function declaration(script, name) {
  const marker = `function ${name}(`;
  const start = script.indexOf(marker);
  assert.ok(start >= 0, `Missing original function ${name}`);
  const open = script.indexOf('{', start);
  let depth = 0, quote = '', comment = '';
  for (let i = open; i < script.length; i++) {
    const c = script[i], next = script[i + 1];
    if (comment === 'line') { if (c === '\n') comment = ''; continue; }
    if (comment === 'block') { if (c === '*' && next === '/') { comment = ''; i++; } continue; }
    if (quote) { if (c === '\\') i++; else if (c === quote) quote = ''; continue; }
    if (c === '/' && next === '/') { comment = 'line'; i++; continue; }
    if (c === '/' && next === '*') { comment = 'block'; i++; continue; }
    if (c === '"' || c === "'" || c === '`') { quote = c; continue; }
    if (c === '{') depth++;
    if (c === '}' && --depth === 0) return script.slice(start, i + 1);
  }
  throw new Error(`Unclosed original function ${name}`);
}
const plain = value => JSON.parse(JSON.stringify(value));

test('chart presentation keeps original series and formatting intact and honors reduced motion', () => {
  const script = fs.readFileSync(new URL('../public/overview/chart-presentation.js', import.meta.url), 'utf8');
  for (const reduced of [false, true]) {
    let plugin;
    const context = vm.createContext({ URLSearchParams, location: { search: '' }, document: { documentElement: { dataset: { theme: 'dark' } } }, matchMedia: () => ({ matches: reduced, addEventListener() {} }), getComputedStyle: () => ({ getPropertyValue: () => '#a4b3ca' }), Chart: { register(value) { plugin = value; }, defaults: { font: {} }, instances: {} }, window: { addEventListener() {} } });
    context.window.Chart = context.Chart;
    vm.runInContext(script, context);
    const formatter = value => String(value), color = () => '#15803d';
    const chart = { data: { labels: ['01/10', '02/10'], datasets: [{ label: 'Receita', data: [0, 1234.56], borderColor: '#2563eb' }, { label: 'Projeção', data: [null, 4321] }] }, config: { options: { animation: { duration: 700 }, plugins: { tooltip: { callbacks: { label: formatter } }, datalabels: { color, formatter } }, scales: { y: { ticks: { callback: formatter }, min: 0 } } } } };
    const data = plain(chart.data);
    plugin.beforeInit(chart); plugin.beforeUpdate(chart);
    assert.deepEqual(plain(chart.data), data);
    assert.equal(chart.config.options.plugins.tooltip.callbacks.label, formatter);
    assert.equal(chart.config.options.plugins.datalabels.formatter, formatter);
    assert.equal(chart.config.options.scales.y.ticks.callback, formatter);
    assert.equal(chart.config.options.scales.y.min, 0);
    assert.equal(reduced ? chart.config.options.animation : chart.config.options.animation.duration, reduced ? false : 420);
  }
});

test('chart tables preserve zero, missing values, projections and the original value formatter', () => {
  const bridge = fs.readFileSync(new URL('../public/overview/bridge.js', import.meta.url), 'utf8');
  let payload;
  const context = vm.createContext({ emit(value) { payload = value; } });
  vm.runInContext(declaration(bridge, 'chartTable'), context);
  const chart = { canvas: { closest: () => ({ querySelector: () => ({ textContent: ' Receita e investimento ' }) }) }, data: { labels: ['01/10', '02/10'], datasets: [{ label: 'Receita', data: [0, 1234.56] }, { label: 'Projeção', data: [null, 4321] }] }, config: { options: { plugins: { tooltip: { callbacks: { label: ({ dataset, raw }) => `${dataset.label}: R$ ${raw.toFixed(2)}` } } } } } };
  context.chartTable(chart);
  assert.deepEqual(plain(payload.chartTable), { title: 'Receita e investimento', headers: ['Período', 'Receita', 'Projeção'], rows: [['01/10', 'R$ 0.00', '—'], ['02/10', 'R$ 1234.56', 'R$ 4321.00']] });
});
function runtime(script) {
  const names = ['BASE_CAMPS', 'MONTHLY', 'TARGET_CPA', 'TARGET_UPSELL', 'MONTHLY_BUDGET', 'MONTHLY_REVENUE_TARGET', 'MONTHLY_PROFIT_TARGET', 'MONTHLY_NET_PROFIT_TARGET', 'TAX_PCT_DEFAULT'];
  const constants = names.map(name => {
    const match = script.match(new RegExp(`\\bconst ${name}\\s*=[\\s\\S]*?;`));
    assert.ok(match, `Missing original constant ${name}`); return match[0];
  });
  const functions = ['targetFor', 'activeTargets', 'taxPctFor', 'netFactor', 'totals', 'prevTotals', 'pctDelta', 'applyCustomDate'].map(name => declaration(script, name));
  const elements = { dateFrom: { value: '' }, dateTo: { value: '' }, dateLabel: { textContent: '' }, dateDD: { classList: { remove() {} } } };
  const context = vm.createContext({ document: { getElementById: id => elements[id] }, sessionStorage: { setItem() {} }, fmt2: value => value, render() {} });
  vm.runInContext(`${constants.join('\n')}\nlet ALL_DAYS=[], filtered=[], MONTHLY_TARGETS=[], activeCamp='', usingRealData=true;\n${functions.join('\n')}\nfunction setFixture(value) { ALL_DAYS=value.days; filtered=ALL_DAYS.filter(d=>d.dateStr>=value.from&&d.dateStr<=value.to); MONTHLY_TARGETS=value.targets; activeCamp=value.campaign||''; usingRealData=value.real!==false; }\nfunction customCount() { return filtered.length; }`, context, { timeout: 1000 });
  return { context, elements, set: fixture => vm.runInContext(`setFixture(${JSON.stringify(fixture)})`, context, { timeout: 1000 }) };
}
const day = (dateStr, revenue, cost, costBureau, conv, gadsConv, checkouts, pedidosUpsell, tax, refundCount = 0, refundValue = 0) => ({ dateStr, revenue, cost, costBureau, costTotal: cost + costBureau, profit: revenue - cost - costBureau, netProfit: +(revenue * (1 - tax / 100) - cost - costBureau).toFixed(2), conv, gadsConv, checkouts, pedidosUpsell, upsellRate: conv ? pedidosUpsell / conv * 100 : 0, refundCount, refundValue });
const fixture = {
  from: '2026-10-01', to: '2026-10-03', targets: [
    { month: '2026-09-01', target_cpa: 11, target_upsell_pct: 31, monthly_budget: 70000, revenue_target: 120000, tax_pct: 10 },
    { month: '2026-10-01', target_cpa: null, target_upsell_pct: 36, tax_pct: 8 },
  ], days: [
    day('2026-09-25', 10000, 5000, 3000, 100, 100, 1000, 20, 10),
    day('2026-09-28', 100, 10, 20, 1, 1, 10, 0, 10),
    day('2026-09-29', 80, 15, 15, 2, 1, 10, 1, 10),
    day('2026-09-30', 120, 20, 30, 2, 1, 20, 1, 10),
    day('2026-10-01', 200, 40, 20, 2, 1, 10, 1, 8, 1, 12),
    day('2026-10-03', 300, 80, 40, 3, 2, 20, 1, 8),
  ],
};

test('the complete loading/calculation script is preserved byte for byte', () => {
  assert.ok(upstream && upstream.length > 50000, 'Original calculation script must be present');
  assert.ok(embedded === upstream, 'Presentation adaptation must not change calculation/loading JavaScript (comparison intentionally omits source/credentials from failure output)');
});

test('all 15 original KPI IDs, labels and overview panels remain present', () => {
  const expected = [
    ['r', 'Receita Bruta'], ['ct', 'Custo Total COGS'], ['p', 'Lucro Bruto'], ['npo', 'Lucro Líquido'], ['eb', 'EBITDA'],
    ['ro', 'ROAS'], ['cpv', 'CPV'], ['ca', 'CAC'], ['ga', 'Gasto Google Ads'], ['gb', 'Gasto Bureau'],
    ['m', 'Margem'], ['cv', 'Taxa de Conversão'], ['up', 'Taxa de Upsell'], ['tr', 'Transações'], ['tk', 'Ticket Médio'],
  ];
  const kpis = script => [...declaration(script, 'buildKPIs').matchAll(/id:'([^']+)'\s*,\s*label:'([^']+)'/g)].map(match => [match[1], match[2]]);
  assert.deepEqual(kpis(upstream), expected);
  assert.deepEqual(kpis(embedded), expected);
  for (const id of ['view-overview', 'todayPace', 'alertBar', 'kpiRow1', 'kpiRow2', 'kpiRow3', 'produtoTesteCard', 'produtoAtualCard', 'adsBalanceCard', 'bureauCard', 'bureauTypeCardGrid', 'refundSection', 'gaugeRow', 'funnelWrap', 'revenueChart', 'profitChart', 'netProfitChart', 'cacChart', 'upsellChart', 'dowGrid', 'convChart', 'campaignFilter', 'dateFrom', 'dateTo']) assert.match(generated, new RegExp(`id=["']${id}["']`));
  for (const dependency of ['chart.umd.min.js', 'chartjs-plugin-datalabels.min.js']) assert.ok(original.includes(dependency) && generated.includes(dependency));
});

test('the twelve original read sources and their loading/fallback implementation remain unchanged', () => {
  const expected = ['revenue_daily', 'upsell_daily', 'upsell_by_type', 'bureau_daily', 'google_ads_campaign_daily', 'refunds_daily', 'ads_balance_history', 'bureau_by_type_daily', 'checkouts_daily', 'monthly_targets', 'revenue_daily_completa', 'checkouts_campaign_daily'];
  const sources = script => [...declaration(script, 'loadSupabaseData').matchAll(/supaGet\(`?['"]?\/rest\/v1\/([a-z_]+)/g)].map(match => match[1]);
  assert.deepEqual(sources(upstream), expected);
  assert.deepEqual(sources(embedded), expected);
  for (const name of ['loadSupabaseData', 'supaGet', 'makeDays', 'targetFor', 'mergeCostIntoAllDays', 'initDashboard']) assert.equal(declaration(embedded, name), declaration(upstream, name));
});

test('runtime totals, monthly inheritance, tax and previous calendar window match the original', () => {
  const originalRuntime = runtime(upstream), generatedRuntime = runtime(embedded), modelRuntime = runtime(readModel().calculationScript);
  for (const engine of [originalRuntime, generatedRuntime, modelRuntime]) {
    engine.set(fixture);
    const total = plain(engine.context.totals()), previous = plain(engine.context.prevTotals());
    assert.equal(total.rev, 500); assert.equal(total.cost, 120); assert.equal(total.costBureau, 60); assert.equal(total.costTotal, 180);
    assert.equal(total.pft, 320); assert.equal(total.netPftOpt, 280); assert.equal(total.conv, 5);
    assert.equal(total.cac, 24); assert.equal(total.cacAds, 40); assert.equal(total.cpv, 36); assert.equal(total.roas, 2.78);
    assert.equal(total.margin, .64); assert.equal(total.ticket, 100); assert.equal(total.upsell, 40);
    assert.ok(Math.abs(total.convRate - 100 / 6) < 1e-10);
    assert.equal(total.refundCount, 1); assert.equal(total.refundValue, 12); assert.equal(total.refundRate, 16.67);
    assert.equal(previous.rev, 300, 'Previous period includes Sep 28–30, three calendar days despite a gap in the current data');
    assert.equal(previous.pft, 190); assert.equal(previous.netPftOpt, 160); assert.equal(previous.conv, 5); assert.equal(previous.cac, 9);
    const target = plain(engine.context.targetFor('2026-10-03'));
    assert.equal(target.cpa, 11); assert.equal(target.upsell, 36); assert.equal(target.budget, 70000); assert.equal(target.revenue, 120000); assert.equal(target.profit, 36000); assert.equal(target.netProfit, 30000); assert.equal(target.taxPct, 8);
    assert.equal(engine.context.targetFor('2026-08-01').cpa, 9);
    assert.ok(Math.abs(engine.context.netFactor('2026-09-30') - .9) < 1e-12);
    assert.ok(Math.abs(engine.context.netFactor('2026-10-01') - .92) < 1e-12);
    assert.deepEqual(plain(engine.context.pctDelta(110, 100, true)), { txt: '↑10.0% vs ant.', cls: 'neg' });
    assert.equal(engine.context.pctDelta(2, 0, false), null);
    engine.set({ ...fixture, campaign: 'RISCO_DOC' });
    assert.deepEqual(plain(engine.context.totals()), total, 'Real campaign selection must not alter global overview totals');
    engine.set({ ...fixture, targets: [...fixture.targets, { month: '2026-11-01', target_cpa: 0, tax_pct: 0 }] });
    assert.equal(engine.context.targetFor('2026-11-01').cpa, 0, 'Explicit zero is not a missing target');
    assert.equal(engine.context.netFactor('2026-11-01'), 1);
  }
  originalRuntime.set(fixture); generatedRuntime.set(fixture); modelRuntime.set(fixture);
  assert.deepEqual(plain(generatedRuntime.context.totals()), plain(originalRuntime.context.totals()));
  assert.deepEqual(plain(generatedRuntime.context.prevTotals()), plain(originalRuntime.context.prevTotals()));
  assert.deepEqual(plain(modelRuntime.context.totals()), plain(originalRuntime.context.totals()));
  assert.deepEqual(plain(modelRuntime.context.prevTotals()), plain(originalRuntime.context.prevTotals()));
});

test('custom ranges longer than 31 days remain supported without truncation', () => {
  const days = Array.from({ length: 64 }, (_, i) => ({ dateStr: new Date(Date.UTC(2026, 0, 1 + i)).toISOString().slice(0, 10) }));
  for (const script of [upstream, embedded, readModel().calculationScript]) {
    const engine = runtime(script);
    engine.set({ days, targets: [], from: '2026-01-01', to: '2026-03-31' });
    engine.elements.dateFrom.value = '2026-01-01'; engine.elements.dateTo.value = '2026-03-31';
    engine.context.applyCustomDate();
    assert.equal(engine.context.customCount(), 64);
  }
});

test('the isolated read-only guard runs before source code and blocks mutations and unrelated endpoints', async () => {
  assert.ok(generated.indexOf('src="./read-only.js"') > 0 && generated.indexOf('src="./read-only.js"') < generated.indexOf(upstream));
  const calls = [], context = vm.createContext({ URL, Request, location: { href: 'https://demo.local/overview/' }, fetch: async (...args) => { calls.push(args); return { ok: true }; } });
  context.window = context;
  const guard = fs.readFileSync(new URL('../public/overview/read-only.js', import.meta.url), 'utf8');
  vm.runInContext(guard, context, { timeout: 1000 });
  const principal = 'https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/revenue_daily';
  const site = 'https://ozquoloetuzynnyzkado.supabase.co/rest/v1/checkouts_daily';
  assert.equal((await context.fetch(principal)).ok, true);
  assert.equal((await context.fetch(site, { method: 'GET' })).ok, true);
  assert.equal((await context.fetch(new Request(principal))).ok, true);
  const successfulReads = calls.length;
  for (const method of ['POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS', 'HEAD']) await assert.rejects(() => context.fetch(principal, { method }), /somente leitura/);
  await assert.rejects(() => context.fetch(new Request(principal, { method: 'POST' })), /somente leitura/);
  await assert.rejects(() => context.fetch('https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/monthly_targets?on_conflict=month', { method: 'POST' }), /somente leitura/);
  for (const url of ['https://example.com/rest/v1/revenue_daily', 'https://ftmgmfdqdqxboiktxcoj.supabase.co/auth/v1/token?grant_type=password', 'https://ftmgmfdqdqxboiktxcoj.supabase.co.evil.example/rest/v1/revenue_daily']) await assert.rejects(() => context.fetch(url), /somente leitura/);
  assert.equal(calls.length, successfulReads, 'Rejected operations must never reach the underlying fetch');
});


test('the same-document model preserves the complete calculation script and original KPI/source contracts', () => {
  const model = readModel();
  assert.equal(typeof model.body, 'string');
  assert.equal(typeof model.calculationScript, 'string');
  assert.ok(model.calculationScript === upstream, 'Shadow DOM adaptation must retain the entire original calculation/loading script (source omitted from failure output)');
  const styles = Array.isArray(model.styles) ? model.styles.join('\n') : model.styles;
  assert.ok(typeof styles === 'string' && styles.length > 1000, 'The original overview needs its scoped presentation styles');
  for (const selector of ['.kpi-card', '.monitor-grid', '.dow-grid']) assert.ok(styles.includes(selector), `Missing style ${selector}`);
  const kpis = script => [...declaration(script, 'buildKPIs').matchAll(/id:'([^']+)'\s*,\s*label:'([^']+)'/g)].map(match => [match[1], match[2]]);
  const actual = kpis(model.calculationScript);
  assert.equal(actual.length, 15);
  assert.deepEqual(actual.map(([id]) => id), ['r', 'ct', 'p', 'npo', 'eb', 'ro', 'cpv', 'ca', 'ga', 'gb', 'm', 'cv', 'up', 'tr', 'tk']);
  assert.deepEqual(actual, kpis(upstream));
  const expectedSources = ['revenue_daily', 'upsell_daily', 'upsell_by_type', 'bureau_daily', 'google_ads_campaign_daily', 'refunds_daily', 'ads_balance_history', 'bureau_by_type_daily', 'checkouts_daily', 'monthly_targets', 'revenue_daily_completa', 'checkouts_campaign_daily'];
  const sources = [...declaration(model.calculationScript, 'loadSupabaseData').matchAll(/supaGet\(`?['"]?\/rest\/v1\/([a-z_]+)/g)].map(match => match[1]);
  assert.deepEqual(sources, expectedSources);
});

test('the same-document model retains visible panels, filters and hidden builder dependencies', () => {
  const { body } = readModel();
  assert.ok(!/<iframe\b/i.test(body), 'The overview model must render in the current document without nested iframes');
  for (const id of ['view-overview', 'todayPace', 'alertBar', 'kpiRow1', 'kpiRow2', 'kpiRow3', 'produtoTesteCard', 'produtoAtualCard', 'adsBalanceCard', 'bureauCard', 'bureauTypeCardGrid', 'refundSection', 'gaugeRow', 'funnelWrap', 'revenueChart', 'profitChart', 'netProfitChart', 'cacChart', 'upsellChart', 'dowGrid', 'convChart', 'campaignFilter', 'dateFrom', 'dateTo', 'dateDD', 'dpWrapper', 'dateLabel', 'freshDot', 'lastUpdated', 'footerBar', 'campaignTable', 'lostIsTable', 'upsellTypeTable', 'metasTable', 'mtMonth', 'metasLoginBox', 'metasSessionBox', 'metasSaveBtn']) assert.ok(new RegExp(`id=["']${id}["']`).test(body), `Original builder requires scoped element ${id}`);
});

test('the dedicated same-document reader permits source GET requests without changing global fetch', async () => {
  const source = fs.readFileSync(new URL('../src/domain/overview.ts', import.meta.url), 'utf8');
  const javascript = ts.transpileModule(source, { compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS } }).outputText;
  const calls = [];
  const globalFetch = async (...args) => { calls.push(args); return { ok: true }; };
  const context = vm.createContext({ exports: {}, URL, Request, location: { href: 'https://demo.local/app' }, fetch: globalFetch });
  vm.runInContext(javascript, context, { timeout: 1000 });
  const read = context.exports.readOnlyOverviewFetch;
  assert.equal(typeof read, 'function');
  const principal = 'https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/revenue_daily?select=*';
  const site = 'https://ozquoloetuzynnyzkado.supabase.co/rest/v1/checkouts_daily';
  const controller = new AbortController();
  const headers = { Range: '0-999' };
  assert.equal((await read(principal, { headers }, controller.signal)).ok, true);
  assert.equal(calls[0][0], principal);
  assert.equal(calls[0][1].headers, headers, 'Pagination/authentication headers retain their identity');
  assert.equal(calls[0][1].signal, controller.signal, 'Unmount cancellation reaches the original fetch');
  const siteURL = new URL(site);
  assert.equal((await read(siteURL, { method: 'get' })).ok, true);
  assert.equal(calls[1][0], siteURL);
  const request = new Request(principal, { headers: { Range: '1000-1999' } });
  assert.equal((await read(request)).ok, true);
  assert.equal(calls[2][0], request, 'Request objects preserve their own headers/options');
  const successfulReads = calls.length;
  for (const method of ['POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS', 'HEAD']) await assert.rejects(() => read(principal, { method }), /somente leitura/);
  await assert.rejects(() => read(new Request(principal, { method: 'POST', body: '{}' })), /somente leitura/);
  await assert.rejects(() => read('https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/monthly_targets?on_conflict=month', { method: 'POST', body: '{}' }), /somente leitura/);
  for (const url of [
    'https://example.com/rest/v1/revenue_daily',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co.evil.example/rest/v1/revenue_daily',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co/auth/v1/token?grant_type=password',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co/functions/v1/refresh',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v10/revenue_daily',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/../../auth/v1/token',
    'http://ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/revenue_daily',
    'https://ftmgmfdqdqxboiktxcoj.supabase.co:8443/rest/v1/revenue_daily',
    'https://operator:secret@ftmgmfdqdqxboiktxcoj.supabase.co/rest/v1/revenue_daily',
    '/rest/v1/revenue_daily',
  ]) await assert.rejects(() => read(url), /somente leitura/);
  assert.equal(calls.length, successfulReads, 'Rejected operations never reach network transport');
  assert.equal(context.fetch, globalFetch, 'Loading and using the reader never monkey-patches global fetch');
});

test('Dashboard and ExecutiveSummary render through the shared same-document engine without iframe adapters', () => {
  for (const name of ['Dashboard', 'ExecutiveSummary']) {
    const source = fs.readFileSync(new URL(`../src/${name}.tsx`, import.meta.url), 'utf8');
    assert.ok(!/<iframe\b/i.test(source), `${name} must not render an iframe`);
    assert.ok(!/createElement\(\s*['"]iframe['"]|HTMLIFrameElement|\.contentWindow\b/.test(source), `${name} must not retain iframe-only transport`);
    assert.ok(/OriginalOverview/.test(source), `${name} must share the original financial engine`);
  }
});

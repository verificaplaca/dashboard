import { overviewReadState, overviewUnavailableMessage } from './domain/overview-publication';
import { scopeOverviewThemeStyles } from './domain/overview-style-scope';
import { mobileLayoutQuery } from './lib/responsive';
import { useGlobalLoading } from './Loading';
import { costValue, netMarginPercent } from './domain/overview-series';
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import type { Chart as ChartInstance } from 'chart.js';
import presentationScript from '../public/overview/chart-presentation.js?raw';
import { useUI } from './ui';
import { observeEntrances } from './motion';
import { createOverviewInsights } from './overview-insights-ui';
import type { MetricDetail } from './domain/overview-insights';
import { createOverviewSections } from './overview-sections';
import { enhanceOverviewChart, selectOverviewChartDay } from './overview-chart-enhancements';
import { styleAnalyticalChart } from './analytical-chart-theme';
import { createOverviewPresentation } from './overview-presentation';
import { readOnlyOverviewFetch } from './domain/overview';
import { readOverviewPeriod, validateOverviewPeriod, overviewPresets, type OverviewPeriod } from './domain/overview-period';
import type { OverviewHomeSnapshot } from './domain/overview-home-snapshot';

export interface DayDetail { date: string; revenue: number; tax: number; ads: number; bureau: number; result: number; example: boolean; }
export interface OverviewEvent { snapshot?: OverviewHomeSnapshot; loading?: boolean; error?: string; metricDetail?: MetricDetail; indicatorInfo?: {title:string;description:string}; dayDetail?: DayDetail; action?: 'rules' | 'export'; ready?: boolean; example?: boolean; period?: string; sections?: Record<string, number>; csv?: string; chartTable?: { title: string; headers: string[]; rows: string[][] }; }
export interface OverviewHandle { scrollToSection: (section: string) => void; exportIndicators: () => void; clearSelection: () => void; refresh: () => void; setPeriod: (period:OverviewPeriod) => void; }
type Model = { body: string; styles: string; calculationScript: string };
let modelPromise: Promise<Model> | undefined, instance = 0;
const loadModel = () => modelPromise ||= fetch(`${import.meta.env.BASE_URL}overview/model.json`, {signal: AbortSignal.timeout(30_000)}).then(response => { if (!response.ok) throw new Error('Não foi possível carregar a Visão Geral.'); return response.json(); }).catch(error => { modelPromise = undefined; throw error; });

/** The preserved engine runs against a scoped DOM in the SAME document. */
export default forwardRef<OverviewHandle, { dataOnly?: boolean; summary?: boolean; view?: 'overview' | 'campanhas' | 'upsell'; onEvent?: (event: OverviewEvent) => void }>(function OriginalOverview({ dataOnly = false, summary = false, view = 'overview', onEvent }, ref) {
  const { theme, setTheme } = useUI();
  const host = useRef<HTMLDivElement>(null), controller = useRef<ReturnType<typeof mountOverview> | null>(null);
  const callback = useRef(onEvent), currentTheme = useRef(theme);
  callback.current = onEvent; currentTheme.current = theme;
  const [error, setError] = useState(''), [loading, setLoading] = useState(true), [attempt,setAttempt] = useState(0);
  useGlobalLoading(loading);
  useImperativeHandle(ref, () => ({ scrollToSection: section => controller.current?.scrollToSection(section), exportIndicators: () => controller.current?.exportIndicators(), clearSelection: () => controller.current?.clearSelection(), setPeriod: period => controller.current?.setPeriod(period), refresh: () => { if(controller.current)controller.current.refresh();else setAttempt(n=>n+1); } }), []);
  useEffect(() => {
    let active = true; setLoading(true); setError(''); callback.current?.({loading:true});
    Promise.all([loadModel(), import('chart.js/auto'), import('chartjs-plugin-datalabels')]).then(([model, charts, labels]) => {
      if (!active || !host.current) return;
      controller.current = mountOverview(charts.default, labels.default, host.current, model, summary, currentTheme.current, event => {
        if(!active)return;
        if(event.error){controller.current?.dispose();controller.current=null;setError(event.error);setLoading(false);}
        callback.current?.(event);
      }, setTheme, view, value => { if(active) { setLoading(value); callback.current?.({loading:value}); } });
    }).catch(reason => { if (active) { setError(String(reason.message || reason)); setLoading(false); callback.current?.({error:String(reason.message||reason),loading:false}); } });
    return () => { active = false; controller.current?.dispose(); controller.current = null; };
  }, [summary, view, attempt]);
  useEffect(() => { controller.current?.setTheme(theme); }, [theme]);
  return <div className={`native-overview ${summary ? 'native-summary' : ''}`} aria-hidden={dataOnly||undefined} inert={dataOnly} style={dataOnly?{position:'fixed',left:'-200vw',top:0,width:1200,height:1,overflow:'hidden',visibility:'hidden',pointerEvents:'none'}:undefined}>
    {error&&!dataOnly&&<section className="overview-unavailable" role="status"><p>{overviewUnavailableMessage}</p><button type="button" className="button secondary" onClick={()=>setAttempt(n=>n+1)}>Tentar novamente</button></section>}
    <div ref={host} hidden={!!error} className="native-overview-host" aria-label={summary ? 'Resumo financeiro das fontes originais' : 'Indicadores e gráficos da Visão Geral'} />
  </div>;
});

function mountOverview(Chart: typeof import('chart.js').Chart, ChartDataLabels: unknown, host: HTMLElement, model: Model, summary: boolean, theme: string, emit: (event: OverviewEvent) => void, changeTheme: (theme: string) => void, view: 'overview' | 'campanhas' | 'upsell', reportLoading: (loading: boolean) => void) {
  const root = host.shadowRoot || host.attachShadow({ mode: 'open' });
  host.dataset.theme = theme === 'dark' ? 'dark' : '';
  host.dataset.mode = summary ? 'summary' : 'full';
  host.dataset.view = view;
  const styles = scopeOverviewThemeStyles(model.styles).replace('.view:not(#view-overview)', `.view:not(#view-${view})`).replace(/:root/g, ':host').replace(/:host:host/g, ':host').replace(/\[data-mode="summary"\]/g, ':host([data-mode="summary"])').replace(/\bhtml\b/g, ':host').replace(/\bbody\b/g, '.native-body');
  root.innerHTML = `<style>${styles}\n:host{display:block;position:relative;min-width:0;font-family:Inter,system-ui,sans-serif} .native-body{margin:0;min-width:0;overflow:visible}:host{overflow:visible} .date-dropdown{z-index:100} #loadingOverlay{display:none!important} .sidebar,.sidebar-backdrop{display:none!important}</style><div class="native-body">${model.body}</div>`;
  const body = root.querySelector<HTMLElement>('.native-body')!;
  root.querySelector('header')!.id = 'section-filters';
  root.querySelectorAll<HTMLElement>(`.view:not(#view-${view}),.sidebar`).forEach(element => { element.inert = true; element.setAttribute('aria-hidden', 'true'); });
  root.getElementById('campaignFilter')!.setAttribute('aria-label', 'Campanha');
  root.getElementById('dateFrom')!.setAttribute('aria-label', 'Data inicial');
  root.getElementById('dateTo')!.setAttribute('aria-label', 'Data final');
  let comparePeriod = false;
  let costMode = 'total';
  let showRevenue = true;
  let active = true, api: Record<string, any> = {}, scheduled = false;
  const id = ++instance, abort = new AbortController(), timeouts = new Set<number>(), frames = new Set<number>(), cleanups: (() => void)[] = [], owned = new Set<ChartInstance>();
  const listen = (target: EventTarget, type: string, listener: EventListener, options?: boolean | AddEventListenerOptions) => { target.addEventListener(type, listener, options); cleanups.push(() => target.removeEventListener(type, listener, options)); };
  const scopedDocument = { body, documentElement: host, getElementById: (name: string) => root.getElementById(name), querySelector: (selector: string) => root.querySelector(selector), querySelectorAll: (selector: string) => root.querySelectorAll(selector), createElement: document.createElement.bind(document), addEventListener: (type: string, listener: EventListener, options?: boolean | AddEventListenerOptions) => {
    if (type !== 'keydown') { listen(body, type, listener, options); return; }
    listen(summary ? body : document, type, event => {
      if (document.querySelector('[role=dialog]')) return;
      const target = event.composedPath()[0];
      const nativeEvent = new Proxy(event, { get(value, key) { if (key === 'target') return target; const result = Reflect.get(value, key, value); return typeof result === 'function' ? result.bind(value) : result; } });
      const before = host.dataset.theme;
      listener(nativeEvent);
      if (host.dataset.theme !== before) { changeTheme(host.dataset.theme === 'dark' ? 'dark' : 'light'); window.dispatchEvent(new Event('vp-theme-change')); }
    }, options);
  } };
  const scopedTimeout = (callback: () => void, delay = 0) => { const timer = window.setTimeout(() => { timeouts.delete(timer); if (active) callback(); }, delay); timeouts.add(timer); return timer; };
  const scopedFrame = (callback: FrameRequestCallback) => { const frame = requestAnimationFrame(time => { frames.delete(frame); if (active) callback(time); }); frames.add(frame); return frame; };
  let pendingReads = 0;
  const scopedFetch = async (input: RequestInfo | URL, options: RequestInit = {}) => {
    pendingReads++; reportLoading(true);
    try {
      const response = await readOnlyOverviewFetch(input, options, AbortSignal.any([abort.signal, AbortSignal.timeout(30_000)]));
      // A response's headers may arrive long before its data. Keep the lease
      // through the body transfer without changing the original response.
      await response.clone().arrayBuffer();
      return response;
    } finally { pendingReads--; schedule(); }
  };
  const plugins: any[] = [];
  function ScopedChart(item: any, config: any) {
    if (!active) throw new DOMException('Visão encerrada', 'AbortError');
    if (item instanceof HTMLCanvasElement && item.id === 'revenueChart') {
      if (config.options.plugins?.legend) config.options.plugins.legend.display = false;
      const rows = api.readChartSnapshot?.().current || [];
      config.options.plugins.datalabels.color = (ctx: any) => ctx.datasetIndex === 0 ? '#477c25' : '#b84659';
      config.options.plugins.datalabels.font = {size: 10, weight:'500'};
      config.options.layout.padding = {top:32, right:28, bottom:22, left:12};
      for (const series of config.data.datasets) {
        const revenue = series.label.includes('Receita');
        const cost = series.label === 'Investimento';
        if (revenue || series.label === 'Meta/dia') series.hidden = !showRevenue;
        const projectedCost = series.label === 'Projeção Investimento';
        if (revenue || cost || projectedCost) {
          series.borderColor = revenue ? '#83cf39' : '#fb8792';
          series.backgroundColor = revenue ? 'rgba(164,231,90,.16)' : 'rgba(251,135,146,.12)';
          series.pointBackgroundColor = series.borderColor; series.pointRadius = 0;
          series.borderWidth = 2; series.tension = .28;
        }
        if (cost) {
          series.label = costMode === 'total' ? 'Custos totais' : costMode === 'ads' ? 'Apenas Ads' : 'Apenas Bureau';
          series.data = series.data.map((v: number | null, i: number) => i >= rows.length ? null : costMode === 'total' ? v : costValue(rows[i], costMode));
        }
        if (projectedCost && costMode !== 'total') series.hidden = true;
      }
    }
    if (item instanceof HTMLCanvasElement && item.id.startsWith('ks_')) {
      const key = item.id.slice(3), totals = api.totals?.(), targets = api.activeTargets?.();
      const negative = ['p', 'npo'].includes(key) && config.data.datasets[0].borderColor === '#dc2626' || key === 'npo' && totals?.netPftOpt < targets?.netProfit || key === 'tk' && totals?.ticket < api.prevTotals?.()?.ticket || key === 'ca' && totals?.cac > targets?.cpa || key === 'up' && totals?.upsell < targets?.upsell;
      const neutral = ['ct', 'ro', 'ga', 'gb', 'tr'].includes(key);
      if (key === 'eb' && !summary && view === 'overview') {
        const days = api.readChartSnapshot?.().current || [];
        config.data.labels = days.map((d: any) => d.dateStr);
        config.data.datasets[0].data = days.map((d: any) => netMarginPercent(d.revenue, d.netProfit));
      }
      config.options.scales.x = { ...config.options.scales.x, offset: true };
      config.options.layout = { padding: 0, autoPadding: false };
      const values = config.data.datasets[0].data.filter((v: unknown) => typeof v === 'number' && Number.isFinite(v));
      config.options.scales.y = {display:false, min:Math.min(0,...values), max:Math.max(0,...values) || 1, grace:0};
      config.options.elements = {point:{radius:0,hoverRadius:0}};
      config.type = 'bar';
      config.data.datasets[0] = { ...config.data.datasets[0], backgroundColor: key === 'eb' ? config.data.datasets[0].data.map((v: number | null) => v === null ? 'transparent' : v < 0 ? '#ff8290' : '#a0f644') : negative ? '#ff8290' : neutral ? '#c6c9d1' : '#a0f644', borderColor: 'transparent', borderWidth: 0, borderRadius: 1, barPercentage: .62, categoryPercentage: .85 };
    }
    if (item instanceof HTMLCanvasElement) styleAnalyticalChart(item.id, config);
    if (!summary && view === 'overview' && ['profitChart', 'netProfitChart', 'revenueChart'].includes(item.id)) {
      config.options.onClick = (_event: unknown, elements: { index: number }[]) => {
        const index = elements[0]?.index;
        if (index !== undefined) { const detail = api.readDay?.(index); if (detail) { owned.forEach(chart => selectOverviewChartDay(chart, index)); emit({ dayDetail: detail }); } }
      };
    }
    if (item instanceof HTMLCanvasElement && api.readChartSnapshot && view === 'overview' && !summary) enhanceOverviewChart(item.id, config, { ...api.readChartSnapshot(), compare: comparePeriod, reducedMotion: matchMedia('(prefers-reduced-motion: reduce)').matches });
    if (item.id === 'revenueChart' && !showRevenue) config.data.datasets.forEach((s: any) => { if (s.label === 'Período anterior') s.hidden = true; });
    // Dense value labels remain available through the existing table on phones.
    const labels=config.options.plugins?.datalabels;
    if(labels&&labels.display!==false&&!item.id.startsWith('ks_')){
      const display=labels.display;
      labels.display=(context:unknown)=>matchMedia(mobileLayoutQuery).matches?false:typeof display==='function'?display(context):display??true;
    }
    const chart = new Chart(item, config);
    const destroy = chart.destroy.bind(chart); chart.destroy = () => { owned.delete(chart); destroy(); };
    owned.add(chart); return chart;
  }
  Object.assign(ScopedChart, {
    defaults: Chart.defaults, getChart: Chart.getChart,
    register: (plugin: any) => {
      const wrapped = { ...plugin, id: `${plugin.id}-${id}`, beforeInit: (chart: ChartInstance) => { if (chart.canvas.getRootNode() === root) plugin.beforeInit?.(chart); }, beforeUpdate: (chart: ChartInstance) => { if (chart.canvas.getRootNode() === root) plugin.beforeUpdate?.(chart); } };
      plugins.push(wrapped); Chart.register(wrapped);
    },
  });
  Object.defineProperty(ScopedChart, 'instances', { get: () => Object.fromEntries([...owned].map(chart => [chart.id, chart])) });
  const scopedWindow = new Proxy(window, { get(target, key) {
    if (key === 'Chart') return ScopedChart;
    if (key === 'addEventListener') return (type: string, listener: EventListener, options?: boolean | AddEventListenerOptions) => listen(window, type, listener, options);
    const value = Reflect.get(target, key, target); return typeof value === 'function' ? value.bind(target) : value;
  } });
  const motion = matchMedia('(prefers-reduced-motion: reduce)');
  const presentationMotion = { get matches() { return motion.matches; }, addEventListener: (type: string, listener: EventListener) => listen(motion, type, listener) };
  new Function('window', 'document', 'Chart', 'location', 'matchMedia', presentationScript)(scopedWindow, scopedDocument, ScopedChart, { search: summary ? '?mode=summary' : '' }, () => presentationMotion);
  const functions = [...model.calculationScript.matchAll(/^(?:async )?function (\w+)\(/gm)].map(match => match[1]);
  const quietConsole = Object.fromEntries(['log', 'warn', 'error', 'info'].map(level => [level, (...args: unknown[]) => { if (active) (console as any)[level](...args); }]));
  api = new Function('window', 'document', 'fetch', 'Chart', 'ChartDataLabels', 'setTimeout', 'clearTimeout', 'requestAnimationFrame', 'cancelAnimationFrame', 'console', `${model.calculationScript}\nreturn {${functions.join(',')},charts,sparkRefs,readChartSnapshot:()=>({current:filtered,allDays:ALL_DAYS}),readDay:(i)=>{const d=filtered[i];if(!d)return null;return {date:d.dateStr,revenue:d.revenue,tax:d.revenue*(1-netFactor(d.dateStr)),ads:(d.costTotal||0)-(d.costBureau||0),bureau:d.costBureau||0,result:d.netProfit||0,example:!usingRealData}},readHomeSnapshot:()=>({real:usingRealData,totals:totals(),targets:activeTargets(),days:filtered.map(d=>({date:d.dateStr,revenue:d.revenue}))}),readPresentation:()=>({funnel:[filtered.reduce((s,d)=>s+(d.checkouts||0),0),filtered.reduce((s,d)=>s+(d.conv||0),0),filtered.reduce((s,d)=>s+(d.pedidosUpsell!=null?d.pedidosUpsell:Math.round((d.conv||0)*((d.upsellRate||0)/100))),0)],totals:totals(),previous:prevTotals(),dayCount:filtered.length,upsellOrders:usingRealData?filtered.reduce((s,d)=>s+(d.pedidosUpsell||0),0):null,targets:activeTargets(),real:usingRealData,months:new Set(filtered.map(d=>d.dateStr.slice(0,7))).size,month:filtered.at(-1)?.dateStr.slice(0,7)})};`)(scopedWindow, scopedDocument, scopedFetch, ScopedChart, ChartDataLabels, scopedTimeout, clearTimeout, scopedFrame, cancelAnimationFrame, quietConsole);
  // The frozen renderer assumes a nonempty date range. Keep exact custom
  // selections (including empty ones) without changing its financial formulas.
  const applyCustomDate = api.applyCustomDate;
  api.applyCustomDate = () => {
    try { applyCustomDate(); } catch(error) { if(api.readChartSnapshot().current.length)throw error; }
    if(!api.readChartSnapshot().current.length){
      const from=(root.getElementById('dateFrom') as HTMLInputElement).value,to=(root.getElementById('dateTo') as HTMLInputElement).value;
      const label=`${from.slice(8,10)}/${from.slice(5,7)} – ${to.slice(8,10)}/${to.slice(5,7)}`;
      root.getElementById('dateRangeHint')!.textContent=`(${label})`;
      root.getElementById('chartSub')!.textContent='Sem dados no período';
    }
  };
  // Navigation lived in a separate upstream script; select the shell-owned view
  // directly without changing any calculation or loading function.
  root.querySelectorAll<HTMLElement>('.view').forEach(element => element.classList.toggle('is-hidden', element.id !== `view-${view}`));
  const repaint = () => { if (active) window.dispatchEvent(new Event('vp-theme-change')); };
  const setTheme = (value: string) => { host.dataset.theme = value === 'dark' ? 'dark' : ''; repaint(); };
  setTheme(theme);
  const bound = new WeakSet<Element>();
  const bindControls = () => {
    root.querySelectorAll<HTMLElement>('.table-wrap').forEach(wrapper=>{
      if(wrapper.hasAttribute('tabindex'))return;
      wrapper.tabIndex=0;wrapper.setAttribute('role','region');
      const label=wrapper.closest('.card')?.querySelector('.card-title')?.textContent?.trim()||'Tabela de dados';
      wrapper.setAttribute('aria-label',label+' · deslize para consultar as colunas');
      const hint=document.createElement('p');hint.className='mobile-table-hint';hint.textContent='Deslize a tabela para consultar todas as colunas.';hint.setAttribute('aria-hidden','true');wrapper.before(hint);
    });
    root.querySelectorAll<HTMLElement>('[onclick],[onchange],[oninput]').forEach(element => {
      if (bound.has(element)) return; bound.add(element);
      for (const eventName of ['click', 'change', 'input']) {
        const expression = element.getAttribute(`on${eventName}`);
        if (!expression) continue;
        // Only functions in the frozen engine are exposed; no global inline handlers.
        const names = Object.keys(api).filter(name => typeof api[name] === 'function');
        const invoke = new Function(...names, 'event', expression);
        // Override the inline handler property while preserving its attribute:
        // the original engine reads that attribute to restore date presets.
        (element as any)[`on${eventName}`] = (event: Event) => {
          invoke.call(element, ...names.map(name => api[name]), event);
          if(expression==='toggleDP()')event.stopPropagation();
          if(element.classList.contains('date-preset')&&expression.includes("'custom'")&&matchMedia(mobileLayoutQuery).matches)root.getElementById('customRange')?.scrollIntoView({block:'nearest'});
          if (expression === 'toggleDP()' && matchMedia(mobileLayoutQuery).matches) {
            const picker = root.getElementById('dateDD')!;
            body.appendChild(picker);
            const top=window.innerHeight<500?(document.querySelector('.topbar')?.getBoundingClientRect().bottom||64)+8:element.getBoundingClientRect().bottom+6;
            picker.style.top=top+'px';
            picker.style.maxHeight=Math.max(120,window.innerHeight-top-16)+'px';
          }
        };
      }
      if (element.classList.contains('date-preset')) { element.tabIndex = 0; element.setAttribute('role', 'button'); listen(element, 'keydown', event => { const key = (event as KeyboardEvent).key; if (key === 'Enter' || key === ' ') { event.preventDefault(); element.click(); } }); }
    });
  };
  bindControls();
  listen(root,'click',event=>{
    if(summary||view!=='overview'||!matchMedia(mobileLayoutQuery).matches)return;
    const button=(event.target as Element).closest('.has-tip .tip-icon');
    const card=button?.closest('.has-tip');
    if(!card)return;
    event.preventDefault();event.stopImmediatePropagation();
    emit({indicatorInfo:{title:card.querySelector('.kpi-label,.gauge-lbl,.card-title')?.textContent?.trim()||'Indicador',description:card.getAttribute('data-tip')||card.getAttribute('data-original-tip')||''}});
  },true);
  const chartTable = (chart: ChartInstance) => {
    const title = chart.canvas.closest('.card,.produto-card')?.querySelector('.card-title')?.textContent?.trim() || 'Valores do gráfico';
    const visible = chart.data.datasets.map((series, datasetIndex) => ({series, datasetIndex})).filter(({datasetIndex}) => chart.isDatasetVisible(datasetIndex));
    const headers = ['Período', ...visible.map(({series}) => series.label || 'Série')];
    const rows = chart.data.labels!.map((label, index) => [String(label), ...visible.map(({series, datasetIndex}) => {
      const value = series.data[index]; if (value == null) return '—';
      const formatter = (chart.config as any).options.plugins?.tooltip?.callbacks?.label;
      const text = formatter ? String(formatter({ raw: value, dataset: series, datasetIndex, dataIndex: index, label, formattedValue: String(value) })).trim() : Number(value).toLocaleString('pt-BR', { maximumFractionDigits: 2 });
      const prefix = `${series.label || ''}:`; return text.startsWith(prefix) ? text.slice(prefix.length).trim() : text;
    })]);
    emit({ chartTable: { title, headers, rows } });
  };
  const replayed = new WeakSet<HTMLCanvasElement>();
  const chartObserver = new IntersectionObserver(entries => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      const canvas = entry.target as HTMLCanvasElement, chart = Chart.getChart(canvas);
      if (!chart) continue;
      chartObserver.unobserve(canvas); replayed.add(canvas);
      if (!motion.matches && !summary && (chart.config as any).options.animation !== false) { chart.reset(); chart.update(); }
    }
  }, { threshold: .15 });
  const present = createOverviewPresentation(root, () => api.readPresentation(), summary, view, action => emit({ action }));
  const presentSections = createOverviewSections(root, () => api.readPresentation(), summary, view);
  const presentInsights = createOverviewInsights(root, () => api.readPresentation(), detail => emit({ metricDetail: detail }), summary || view !== 'overview');
  let previous = '';
  const sync = () => {
    scheduled = false; if (!active) return;
    const sourceReady=root.querySelectorAll('.kpi-row .kpi-value').length===15;
    const sourceState=overviewReadState({exampleBanner:!!root.getElementById('demoDataBanner'),ready:sourceReady,real:sourceReady&&api.readHomeSnapshot().real,pendingReads});
    if(sourceState==='unavailable'){
      reportLoading(false);emit({error:overviewUnavailableMessage,ready:false,loading:false});return;
    }
    bindControls(); present(); presentSections(); presentInsights();
    if (!summary) root.querySelectorAll<HTMLCanvasElement>(`#view-${view} .card canvas,#view-${view} .produto-card canvas`).forEach(canvas => {
      canvas.setAttribute('role', 'img');
      const header = canvas.closest('.card,.produto-card')?.querySelector('.product-chart-header,.card-header'); if (!header) return;
      canvas.setAttribute('aria-label', canvas.closest('.produto-card')?.querySelector('.produto-card-label')?.textContent || header.querySelector('.card-title')?.textContent || 'Gráfico de performance');
      let button = header.querySelector<HTMLButtonElement>('.chart-values-button');
      if (!button) { button = document.createElement('button'); button.className = 'chart-values-button'; button.textContent = 'Ver valores'; button.setAttribute('aria-label', `Ver valores de ${canvas.getAttribute('aria-label')}`); button.onclick = () => { const chart = Chart.getChart(canvas); if (chart) chartTable(chart); }; header.appendChild(button); }
      const chart = Chart.getChart(canvas); button.disabled = !chart;
      if (chart && ['revenueChart', 'profitChart', 'netProfitChart'].includes(canvas.id) && !header.querySelector('.chart-comparison-control')) {
        const label = document.createElement('label'); label.className = 'chart-comparison-control';
        const checkbox = document.createElement('input'); checkbox.type = 'checkbox'; checkbox.checked = comparePeriod;
        label.append(checkbox, document.createTextNode('Comparar: mesma duração anterior'));
        checkbox.onchange = () => { comparePeriod = checkbox.checked; root.querySelectorAll<HTMLInputElement>('.chart-comparison-control input').forEach(input => input.checked = comparePeriod); api.buildCharts(api.totals()); schedule(); };
        header.append(label);
        const key = document.createElement('p'); key.className = 'chart-reading-key'; key.textContent = 'Realizado: contínuo · Anterior: cinza tracejado · Projeções e metas: identificadas na legenda'; header.after(key);
      }
      if (chart && canvas.id === 'revenueChart' && !header.querySelector('.cost-switcher')) {
        const group = document.createElement('div'); group.className = 'cost-switcher'; group.setAttribute('role', 'group'); group.setAttribute('aria-label', 'Custos comparados à receita');
        for (const [mode, name] of [['total', 'Custos totais'], ['ads', 'Apenas Ads'], ['bureau', 'Apenas Bureau']]) {
          const button = document.createElement('button'); button.textContent = name; button.setAttribute('aria-pressed', String(costMode === mode));
          button.onclick = () => { costMode = mode; group.querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(b === button))); api.buildCharts(api.totals()); schedule(); };
          group.append(button);
        }
        header.append(group);
        const revenueButton = document.createElement('button'); revenueButton.className = 'revenue-toggle'; revenueButton.textContent = 'Receita'; revenueButton.setAttribute('aria-pressed', String(showRevenue));
        revenueButton.onclick = () => { showRevenue = !showRevenue; revenueButton.setAttribute('aria-pressed', String(showRevenue)); api.buildCharts(api.totals()); schedule(); };
        header.querySelector('.revenue-legend')?.replaceWith(revenueButton);
      }
      if (chart && ['revenueChart', 'profitChart', 'netProfitChart'].includes(canvas.id)) {
        const key = header.parentElement?.querySelector('.chart-reading-key');
        const available = chart.data.datasets.some((d: any) => d.label === 'Período anterior');
        const text = comparePeriod ? available ? 'Anterior: cinza tracejado · Datas equivalentes nos detalhes do gráfico.' : 'Sem dados equivalentes disponíveis no período anterior.' : 'Ative a comparação para consultar o período anterior de mesma duração.';
        if (key && key.textContent !== text) key.textContent = text;
      }
      if (chart && ['netProfitChart', 'profitChart'].includes(canvas.id)) {
        let select = header.querySelector<HTMLSelectElement>('.day-detail-select');
        if (!select) { select = document.createElement('select'); select.className = 'day-detail-select'; select.setAttribute('aria-label', 'Investigar resultado de um dia'); header.appendChild(select); select.onchange = () => { if (select!.value !== '') { const detail = api.readDay(Number(select!.value)); if (detail) { owned.forEach(chart => selectOverviewChartDay(chart, Number(select!.value))); emit({ dayDetail: detail }); } select!.value = ''; } }; }
        const labels = (chart.data.labels || []).map(String);
        if (select.dataset.labels !== JSON.stringify(labels)) { select.replaceChildren(new Option('Investigar dia…', ''), ...labels.map((label, i) => new Option(label, String(i)))); select.dataset.labels = JSON.stringify(labels); }
      }
      if (chart && !replayed.has(canvas)) chartObserver.observe(canvas);
    });
    const event = { ready: root.querySelectorAll('.kpi-row .kpi-value').length === 15, example: false, period: root.getElementById('dateRangeHint')!.textContent!.trim(), sections: Object.fromEntries([...root.querySelectorAll<HTMLElement>('[id^="section-"]')].filter(el => el.getBoundingClientRect().height).map(el => [el.id.replace('section-', ''), el.getBoundingClientRect().top + window.scrollY])) };
    const nativeOverlay = root.getElementById('loadingOverlay');
    const finished = event.ready || !nativeOverlay || nativeOverlay.style.opacity === '0';
    const busy = pendingReads > 0 || !finished;
    if(!busy&&event.ready){
      const selection=readOverviewPeriod(sessionStorage),days=api.readChartSnapshot().current as {dateStr:string}[];
      if(selection.preset==='custom'&&days.some(day=>day.dateStr<selection.from!||day.dateStr>selection.to!)){
        (root.getElementById('dateFrom') as HTMLInputElement).value=selection.from!;
        (root.getElementById('dateTo') as HTMLInputElement).value=selection.to!;
        api.applyCustomDate();schedule();return;
      }
    }
    host.setAttribute('aria-busy', String(busy));
    reportLoading(busy);
    const output: OverviewEvent = {...event,...(!busy && event.ready ? {snapshot:{...api.readHomeSnapshot(),filter:readOverviewPeriod(sessionStorage)}} : {})};
    const key = JSON.stringify(output); if (key !== previous) { previous = key; emit(output); }
  };
  const schedule = () => { if (!scheduled && active) { scheduled = true; scopedFrame(sync); } };
  const observer = new MutationObserver(schedule); observer.observe(root, { subtree: true, childList: true, attributes: true, attributeFilter: ['style', 'class'] });
  const resize = new ResizeObserver(schedule); resize.observe(body);
  const noteParent = root.getElementById('chartSub')?.parentElement;
  if (noteParent) { const note = document.createElement('p'); note.className = 'chart-context-note'; note.textContent = 'Investimento = Google Ads + bureau. Linhas tracejadas indicam projeções e metas.'; noteParent.appendChild(note); }
  const outside = (event: Event) => { const tools = root.querySelector<HTMLDetailsElement>('.overview-tools'); if (tools?.open && !event.composedPath().includes(tools)) tools.open = false; if (!event.composedPath().includes(host)) { root.getElementById('dateDD')?.classList.remove('open'); root.querySelectorAll('.has-tip.tip-open').forEach(el => el.classList.remove('tip-open')); } };
  listen(document, 'click', outside);
  listen(document, 'keydown', event => { if ((event as KeyboardEvent).key === 'Escape') { const tools = root.querySelector<HTMLDetailsElement>('.overview-tools'); if (tools?.open) tools.open = false; const picker=root.getElementById('dateDD');if(picker?.classList.contains('open')){picker.classList.remove('open');root.querySelector<HTMLButtonElement>('.date-picker-btn')?.focus();} } });

  const stopEntrances = observeEntrances(root); schedule();
  return {
    setTheme,
    refresh: () => { void api.initDashboard(); },
    setPeriod: (selection:OverviewPeriod) => {
      validateOverviewPeriod(selection);
      if(selection.preset==='custom'){
        (root.getElementById('dateFrom') as HTMLInputElement).value=selection.from!;
        (root.getElementById('dateTo') as HTMLInputElement).value=selection.to!;
        api.applyCustomDate();
      }else{
        const currentTarget=[...root.querySelectorAll<HTMLElement>('.date-preset')].find(element=>element.getAttribute('onclick')?.includes("'"+selection.preset+"'"));
        if(!currentTarget)throw new Error('Filtro indisponível. Atualize o resumo.');
        api.setPreset(selection.preset,overviewPresets.find(([key])=>key===selection.preset)![1],{currentTarget});
      }
      schedule();
    },
    clearSelection: () => owned.forEach(chart => selectOverviewChartDay(chart, -1)),
    scrollToSection: (section: string) => { const element = root.getElementById(`section-${section}`); if (element) window.scrollTo({ top: element.getBoundingClientRect().top + window.scrollY - (document.querySelector('.topbar')?.getBoundingClientRect().height || 72) - (document.querySelector('.overview-navigation')?.getBoundingClientRect().height || 0) - 14, behavior: motion.matches ? 'instant' : 'smooth' }); },
    exportIndicators: () => { const quote = (value: unknown) => '"' + String(value ?? '').replace(/"/g, '""') + '"'; const rows = [...root.querySelectorAll('.kpi-row .kpi-card')].map(card => ['.kpi-label', '.kpi-value', '.kpi-value-secondary', '.kpi-delta', '.kpi-sub'].map(selector => quote(selector === '.kpi-delta' ? card.querySelector(selector)?.getAttribute('data-original-delta') ?? card.querySelector(selector)?.textContent : card.querySelector(selector)?.textContent)).join(';')); emit({ csv: ['Período;' + quote(root.getElementById('dateRangeHint')?.textContent), 'Fonte;' + quote(root.getElementById('footerBar')?.textContent), 'Indicador;Valor;Complemento;Comparativo;Contexto', ...rows].join('\n') }); },
    dispose: () => { if(!active)return;active = false; abort.abort(); observer.disconnect(); resize.disconnect(); chartObserver.disconnect(); stopEntrances(); presentSections.dispose?.(); cleanups.forEach(cleanup => cleanup()); timeouts.forEach(clearTimeout); frames.forEach(cancelAnimationFrame); owned.forEach(chart => chart.destroy()); plugins.forEach(plugin => Chart.unregister(plugin)); root.replaceChildren(); },
  };
}

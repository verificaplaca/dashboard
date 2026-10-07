import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import type { Chart as ChartInstance } from 'chart.js';
import presentationScript from '../public/overview/chart-presentation.js?raw';
import { useUI } from './ui';
import { observeEntrances } from './motion';
import { readOnlyOverviewFetch } from './domain/overview';

export interface OverviewEvent { ready?: boolean; example?: boolean; period?: string; sections?: Record<string, number>; csv?: string; chartTable?: { title: string; headers: string[]; rows: string[][] }; }
export interface OverviewHandle { scrollToSection: (section: string) => void; exportIndicators: () => void; }
type Model = { body: string; styles: string; calculationScript: string };
let modelPromise: Promise<Model> | undefined, instance = 0;
const loadModel = () => modelPromise ||= fetch(`${import.meta.env.BASE_URL}overview/model.json`).then(response => { if (!response.ok) throw new Error('Não foi possível carregar a Visão Geral.'); return response.json(); }).catch(error => { modelPromise = undefined; throw error; });

/** The preserved engine runs against a scoped DOM in the SAME document. */
export default forwardRef<OverviewHandle, { summary?: boolean; view?: 'overview' | 'campanhas' | 'upsell'; onEvent?: (event: OverviewEvent) => void }>(function OriginalOverview({ summary = false, view = 'overview', onEvent }, ref) {
  const { theme, setTheme } = useUI();
  const host = useRef<HTMLDivElement>(null), controller = useRef<ReturnType<typeof mountOverview> | null>(null);
  const callback = useRef(onEvent), currentTheme = useRef(theme);
  callback.current = onEvent; currentTheme.current = theme;
  const [error, setError] = useState('');
  useImperativeHandle(ref, () => ({ scrollToSection: section => controller.current?.scrollToSection(section), exportIndicators: () => controller.current?.exportIndicators() }), []);
  useEffect(() => {
    let active = true;
    Promise.all([loadModel(), import('chart.js/auto'), import('chartjs-plugin-datalabels')]).then(([model, charts, labels]) => {
      if (!active || !host.current) return;
      controller.current = mountOverview(charts.default, labels.default, host.current, model, summary, currentTheme.current, event => callback.current?.(event), setTheme, view);
    }).catch(reason => { if (active) setError(String(reason.message || reason)); });
    return () => { active = false; controller.current?.dispose(); controller.current = null; };
  }, [summary, view]);
  useEffect(() => { controller.current?.setTheme(theme); }, [theme]);
  return <div className={`native-overview ${summary ? 'native-summary' : ''}`}>
    {error && <p role="alert" className="info-note">{error} Recarregue a página para tentar novamente.</p>}
    <div ref={host} className="native-overview-host" aria-label={summary ? 'Resumo financeiro das fontes originais' : 'Indicadores e gráficos da Visão Geral'} />
  </div>;
});

function mountOverview(Chart: typeof import('chart.js').Chart, ChartDataLabels: unknown, host: HTMLElement, model: Model, summary: boolean, theme: string, emit: (event: OverviewEvent) => void, changeTheme: (theme: string) => void, view: 'overview' | 'campanhas' | 'upsell') {
  const root = host.shadowRoot || host.attachShadow({ mode: 'open' });
  host.dataset.theme = theme === 'dark' ? 'dark' : '';
  host.dataset.mode = summary ? 'summary' : 'full';
  const styles = model.styles.replace('.view:not(#view-overview)', `.view:not(#view-${view})`).replace(/:root/g, ':host').replace(/\[data-theme="dark"\]/g, ':host([data-theme="dark"])').replace(/:host:host/g, ':host').replace(/\[data-mode="summary"\]/g, ':host([data-mode="summary"])').replace(/\bhtml\b/g, ':host').replace(/\bbody\b/g, '.native-body');
  root.innerHTML = `<style>${styles}\n:host{display:block;position:relative;min-width:0;font-family:Inter,system-ui,sans-serif} .native-body{margin:0;min-width:0;overflow:visible}:host{overflow:visible} .date-dropdown{z-index:100} .kpi-row:empty::before{content:'Carregando indicadores…';padding:22px;color:var(--text-secondary);font-size:12px} .sidebar,.sidebar-backdrop{display:none!important}</style><div class="native-body">${model.body}</div>`;
  const body = root.querySelector<HTMLElement>('.native-body')!;
  root.querySelector('header')!.id = 'section-filters';
  root.querySelectorAll<HTMLElement>(`.view:not(#view-${view}),.sidebar`).forEach(element => { element.inert = true; element.setAttribute('aria-hidden', 'true'); });
  root.getElementById('campaignFilter')!.setAttribute('aria-label', 'Campanha');
  root.getElementById('dateFrom')!.setAttribute('aria-label', 'Data inicial');
  root.getElementById('dateTo')!.setAttribute('aria-label', 'Data final');
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
  const scopedFetch = (input: RequestInfo | URL, options: RequestInit = {}) => readOnlyOverviewFetch(input, options, abort.signal);
  const plugins: any[] = [];
  function ScopedChart(item: any, config: any) {
    if (!active) throw new DOMException('Visão encerrada', 'AbortError');
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
  api = new Function('window', 'document', 'fetch', 'Chart', 'ChartDataLabels', 'setTimeout', 'clearTimeout', 'requestAnimationFrame', 'cancelAnimationFrame', 'console', `${model.calculationScript}\nreturn {${functions.join(',')},charts,sparkRefs};`)(scopedWindow, scopedDocument, scopedFetch, ScopedChart, ChartDataLabels, scopedTimeout, clearTimeout, scopedFrame, cancelAnimationFrame, quietConsole);
  // Navigation lived in a separate upstream script; select the shell-owned view
  // directly without changing any calculation or loading function.
  root.querySelectorAll<HTMLElement>('.view').forEach(element => element.classList.toggle('is-hidden', element.id !== `view-${view}`));
  const repaint = () => { if (active) window.dispatchEvent(new Event('vp-theme-change')); };
  const setTheme = (value: string) => { host.dataset.theme = value === 'dark' ? 'dark' : ''; repaint(); };
  setTheme(theme);
  const bound = new WeakSet<Element>();
  const bindControls = () => {
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
          if (expression === 'toggleDP()' && window.innerWidth <= 600) {
            const picker = root.getElementById('dateDD')!;
            picker.style.maxHeight = Math.max(180, window.innerHeight - picker.getBoundingClientRect().top - 16) + 'px';
          }
        };
      }
      if (element.classList.contains('date-preset')) { element.tabIndex = 0; element.setAttribute('role', 'button'); listen(element, 'keydown', event => { const key = (event as KeyboardEvent).key; if (key === 'Enter' || key === ' ') { event.preventDefault(); element.click(); } }); }
    });
  };
  bindControls();
  const chartTable = (chart: ChartInstance) => {
    const title = chart.canvas.closest('.card')?.querySelector('.card-title')?.textContent?.trim() || 'Valores do gráfico';
    const headers = ['Período', ...chart.data.datasets.map(series => series.label || 'Série')];
    const rows = chart.data.labels!.map((label, index) => [String(label), ...chart.data.datasets.map((series, datasetIndex) => {
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
  let previous = '';
  const sync = () => {
    scheduled = false; if (!active) return; bindControls();
    if (!summary) root.querySelectorAll<HTMLCanvasElement>('#view-overview .card canvas').forEach(canvas => {
      canvas.setAttribute('role', 'img');
      const header = canvas.closest('.card')?.querySelector('.card-header'); if (!header) return;
      canvas.setAttribute('aria-label', header.querySelector('.card-title')?.textContent || 'Gráfico de performance');
      let button = header.querySelector<HTMLButtonElement>('.chart-values-button');
      if (!button) { button = document.createElement('button'); button.className = 'chart-values-button'; button.textContent = 'Ver valores'; button.setAttribute('aria-label', `Ver valores de ${canvas.getAttribute('aria-label')}`); button.onclick = () => { const chart = Chart.getChart(canvas); if (chart) chartTable(chart); }; header.appendChild(button); }
      const chart = Chart.getChart(canvas); button.disabled = !chart;
      if (chart && !replayed.has(canvas)) chartObserver.observe(canvas);
    });
    const event = { ready: root.querySelectorAll('.kpi-row .kpi-value').length === 15, example: !!root.getElementById('demoDataBanner'), period: root.getElementById('dateRangeHint')!.textContent!.trim(), sections: Object.fromEntries([...root.querySelectorAll<HTMLElement>('[id^="section-"]')].filter(el => el.getBoundingClientRect().height).map(el => [el.id.replace('section-', ''), el.getBoundingClientRect().top + window.scrollY])) };
    const key = JSON.stringify(event); if (key !== previous) { previous = key; emit(event); }
  };
  const schedule = () => { if (!scheduled && active) { scheduled = true; scopedFrame(sync); } };
  const observer = new MutationObserver(schedule); observer.observe(root, { subtree: true, childList: true, attributes: true, attributeFilter: ['style', 'class'] });
  const resize = new ResizeObserver(schedule); resize.observe(body);
  const noteParent = root.getElementById('chartSub')?.parentElement;
  if (noteParent) { const note = document.createElement('p'); note.className = 'chart-context-note'; note.textContent = 'Investimento = Google Ads + bureau. Linhas tracejadas indicam projeções e metas.'; noteParent.appendChild(note); }
  const outside = (event: Event) => { if (!event.composedPath().includes(host)) { root.getElementById('dateDD')?.classList.remove('open'); root.querySelectorAll('.has-tip.tip-open').forEach(el => el.classList.remove('tip-open')); } };
  listen(document, 'click', outside);

  const stopEntrances = observeEntrances(root); schedule();
  return {
    setTheme,
    scrollToSection: (section: string) => { const element = root.getElementById(`section-${section}`); if (element) window.scrollTo({ top: element.getBoundingClientRect().top + window.scrollY - (document.querySelector('.topbar')?.getBoundingClientRect().height || 72) - (document.querySelector('.overview-navigation')?.getBoundingClientRect().height || 0) - 14, behavior: motion.matches ? 'instant' : 'smooth' }); },
    exportIndicators: () => { const quote = (value: unknown) => '"' + String(value ?? '').replace(/"/g, '""') + '"'; const rows = [...root.querySelectorAll('.kpi-row .kpi-card')].map(card => ['.kpi-label', '.kpi-value', '.kpi-value-secondary', '.kpi-delta', '.kpi-sub'].map(selector => quote(card.querySelector(selector)?.textContent)).join(';')); emit({ csv: ['Período;' + quote(root.getElementById('dateRangeHint')?.textContent), 'Fonte;' + quote(root.getElementById('footerBar')?.textContent), 'Indicador;Valor;Complemento;Comparativo;Contexto', ...rows].join('\n') }); },
    dispose: () => { active = false; abort.abort(); observer.disconnect(); resize.disconnect(); chartObserver.disconnect(); stopEntrances(); cleanups.forEach(cleanup => cleanup()); timeouts.forEach(clearTimeout); frames.forEach(cancelAnimationFrame); owned.forEach(chart => chart.destroy()); plugins.forEach(plugin => Chart.unregister(plugin)); root.replaceChildren(); },
  };
}

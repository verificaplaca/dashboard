import { chartTooltipCopy } from './domain/chart-tooltip';
import { mobileLayoutQuery } from './lib/responsive';

type TooltipChart = { options: any; canvas: HTMLCanvasElement; tooltip?: any; update: (mode: 'none') => void };
const states = new WeakMap<TooltipChart, { configure: () => void; dispose: () => void }>();

/** Mobile values live outside the canvas so long labels can wrap and scroll. */
export const mobileChartTooltips = {
  id: 'mobile-chart-tooltips',
  beforeInit(chart: TooltipChart) {
    const query = matchMedia(mobileLayoutQuery);
    const original = chart.options.plugins?.tooltip;
    const enabled = original?.enabled ?? true, external = original?.external;
    let panel: HTMLElement | undefined, heading: HTMLElement, content: HTMLElement;
    const hide = () => { if (panel) panel.hidden = true; };
    const dismiss = () => { hide(); chart.tooltip?.setActiveElements([], { x: 0, y: 0 }); };
    const render = ({ tooltip }: any) => {
      if (!query.matches) { hide(); external?.({ chart, tooltip }); return; }
      if (document.querySelector('[role="dialog"]')) { hide(); return; }
      // Keep the last tapped values available when the panel covers the pointer
      // or the finger leaves the canvas. Close explicitly or by tapping outside.
      if (!tooltip.opacity) return;
      if (!panel) {
        panel = document.createElement('section'); panel.className = 'mobile-chart-tooltip';
        panel.setAttribute('role', 'region'); panel.setAttribute('aria-label', 'Valores do gráfico');
        const header = document.createElement('header');
        heading = document.createElement('strong');
        const close = document.createElement('button'); close.type = 'button'; close.textContent = 'Fechar';
        close.setAttribute('aria-label', 'Fechar valores do gráfico'); close.onclick = dismiss;
        header.append(heading, close);
        content = document.createElement('div'); content.className = 'mobile-chart-tooltip-content';
        content.setAttribute('aria-live', 'polite'); content.setAttribute('aria-atomic', 'true');
        panel.append(header, content); document.body.append(panel);
      }
      const copy = chartTooltipCopy(tooltip);
      heading.textContent = copy.title || 'Valores do gráfico';
      const key = JSON.stringify(copy);
      if (panel.dataset.copy !== key) {
        panel.dataset.copy = key; content.replaceChildren(); content.scrollTop = 0;
        for (const line of copy.lines) { const row = document.createElement('p'); row.textContent = line; content.append(row); }
      }
      document.querySelectorAll<HTMLElement>('.mobile-chart-tooltip').forEach(other => { if (other !== panel) other.hidden = true; });
      panel.hidden = false;
    };
    const configure = () => {
      const tooltip = chart.options.plugins.tooltip;
      if (tooltip === false) return;
      tooltip.enabled = query.matches ? false : enabled;
      tooltip.external = render;
    };
    const outside = (event: Event) => { const path = event.composedPath(); if (!path.includes(chart.canvas) && (!panel || !path.includes(panel))) dismiss(); };
    const scroll = (event: Event) => { if (!panel?.contains(event.target as Node)) hide(); };
    const key = (event: KeyboardEvent) => { if (event.key === 'Escape') dismiss(); };
    const resize = () => { hide(); chart.update('none'); };
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('scroll', scroll, true);
    document.addEventListener('keydown', key);
    query.addEventListener('change', resize);
    states.set(chart, { configure, dispose: () => {
      panel?.remove(); document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('scroll', scroll, true); document.removeEventListener('keydown', key);
      query.removeEventListener('change', resize);
    } });
    configure();
  },
  beforeUpdate(chart: TooltipChart) { states.get(chart)?.configure(); },
  afterDestroy(chart: TooltipChart) { states.get(chart)?.dispose(); states.delete(chart); },
};

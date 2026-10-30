import { buildPeriodSummary, buildMetricDetail, type MetricDetail } from './domain/overview-insights';

export function createOverviewInsights(root: ShadowRoot, read: () => any, open: (detail: MetricDetail) => void, disabled: boolean) {
  if (disabled) return () => {};
  const summary = document.createElement('section'); summary.className = 'period-reading'; summary.setAttribute('aria-label', 'Leitura do período');
  summary.innerHTML = '<span class="reading-eyebrow">LEITURA DO PERÍODO</span><p aria-live="polite">Carregando a leitura dos dados…</p><div class="reading-actions"></div>';
  root.querySelector('.section-tabs')?.after(summary);
  const actions = summary.querySelector('.reading-actions')!;
  for (const [key, label] of [['cac', 'Entender CAC'], ['margin', 'Entender margem'], ['upsell', 'Entender upsell']] as const) {
    const button = document.createElement('button'); button.textContent = label + ' ↗'; button.onclick = () => open(buildMetricDetail(key, read())); actions.append(button);
  }
  return () => {
    const snapshot = read(); const value = buildPeriodSummary(snapshot);
    const paragraph = summary.querySelector('p')!;
    if (paragraph.textContent !== value) paragraph.textContent = value;
    for (const [id, key] of [['cacChart', 'cac'], ['upsellChart', 'upsell']] as const) {
      const header = root.getElementById(id)?.closest('.card')?.querySelector('.card-header');
      if (header) {
        const detail = buildMetricDetail(key, snapshot);
        let distance = header.parentElement?.querySelector<HTMLElement>('.metric-target-distance');
        if (!distance) { distance = document.createElement('p'); distance.className = 'metric-target-distance'; header.after(distance); }
        if (distance.textContent !== detail.target) distance.textContent = detail.target;
        if (distance.dataset.tone !== detail.tone) distance.dataset.tone = detail.tone;
      }
      if (header && !header.querySelector('.metric-explain')) {
        const button = document.createElement('button'); button.className = 'metric-explain'; button.textContent = 'Entender indicador ↗'; button.onclick = () => open(buildMetricDetail(key, read())); header.append(button);
      }
    }
  };
}

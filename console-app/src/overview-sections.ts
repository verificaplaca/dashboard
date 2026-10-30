import { createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { FunnelChart } from './components/ui/funnel-chart';
import { resultStatus } from './domain/metric-presentation';

const currency = (value: number) => value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const text = (element: Element | null, value: string) => { if (element && element.textContent !== value) element.textContent = value; };

/** Layout only: original IDs, charts and dynamic builder containers stay intact. */
export function createOverviewSections(root: ShadowRoot, read: () => any, summary: boolean, view: string) {
  if (summary || view !== 'overview') return Object.assign(() => {}, { dispose: () => {} });
  const overview = root.getElementById('view-overview')!;
  const headings: Record<string, [string, string]> = {
    'Por Produto': ['Produtos e conversão', 'Compare a rentabilidade dos produtos e acompanhe a conversão dos pedidos.'],
    'Monitoramento': ['Monitoramento operacional', 'Saldo da conta de mídia e custos de bureau. Cada bloco informa sua janela de leitura.'],
    'Estornos': ['Estornos', 'Valores informativos: cancelamentos já tratados na receita não são descontados novamente.'],
    'Saúde': ['Eficiência e metas', 'Uma leitura do custo de aquisição, margem, adicionais e ritmo do orçamento.'],
    'Funil': ['Funil de conversão', 'Do checkout ao pedido com adicional. Quantidades e taxas da seleção atual.'],
    'Performance Diária': ['Performance diária', 'Resultado e eficiência ao longo do período. Metas e médias mantêm suas referências.'],
    'Performance por Dia': ['O ritmo da semana', 'Médias por dia com dados, de segunda a domingo. O destaque considera o resultado após imposto.'],
  };
  overview.querySelectorAll<HTMLElement>(':scope > .section-label').forEach(label => {
    const match = Object.entries(headings).find(([prefix]) => label.textContent?.trim().startsWith(prefix));
    if (!match) return;
    // Keep nested builder-owned IDs (including the product source/demo chip).
    const firstText = [...label.childNodes].find(node => node.nodeType === Node.TEXT_NODE);
    if (firstText) firstText.textContent = match[1][0];
    label.classList.add('overview-section-heading'); label.dataset.sectionKey = match[0];
    const description = document.createElement('p'); description.className = 'overview-section-description'; description.textContent = match[1][1]; label.after(description);
  });
  const cleanups: (() => void)[] = [];
  const motion = matchMedia('(prefers-reduced-motion: reduce)');
  const sectionNav = document.createElement('nav'); sectionNav.className = 'section-tabs'; sectionNav.setAttribute('aria-label', 'Seções da página');
  const sections = [['financial', 'Resultados'], ['performance', 'Performance'], ['efficiency', 'Eficiência'], ['products', 'Produtos'], ['operation', 'Operação']];
  const navButtons = new Map<string, HTMLButtonElement>();
  const offset = () => (document.querySelector('.topbar')?.getBoundingClientRect().height || 84);
  for (const [id, label] of sections) {
    const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.setAttribute('aria-controls', 'section-' + id);
    button.onclick = () => {
      const target = root.getElementById('section-' + id); if (!target) return;
      window.scrollTo({ top: target.getBoundingClientRect().top + window.scrollY - offset() - sectionNav.offsetHeight - 18, behavior: motion.matches ? 'instant' : 'smooth' });
    };
    navButtons.set(id, button); sectionNav.append(button);
  }
  root.querySelector('.overview-hero')?.after(sectionNav);
  let scrollFrame = 0;
  const updateNavigation = () => {
    scrollFrame = 0;
    const top = offset(), threshold = top + sectionNav.offsetHeight + 30;
    const topValue = top + 'px';
    if (sectionNav.style.getPropertyValue('--section-nav-top') !== topValue) sectionNav.style.setProperty('--section-nav-top', topValue);
    let current = 'financial';
    for (const [id] of sections) { const el = root.getElementById('section-' + id); if (el && el.getBoundingClientRect().top <= threshold) current = id; }
    navButtons.forEach((button, id) => { const state = id === current ? 'location' : 'false'; if (button.getAttribute('aria-current') !== state) button.setAttribute('aria-current', state); });
  };
  const scheduleNavigation = () => { if (!scrollFrame) scrollFrame = requestAnimationFrame(updateNavigation); };
  window.addEventListener('scroll', scheduleNavigation, { passive: true }); window.addEventListener('resize', scheduleNavigation);
  const topbar = document.querySelector('.topbar'); const navResize = new ResizeObserver(scheduleNavigation); if (topbar) navResize.observe(topbar);
  cleanups.push(() => { window.removeEventListener('scroll', scheduleNavigation); window.removeEventListener('resize', scheduleNavigation); navResize.disconnect(); cancelAnimationFrame(scrollFrame); });
  updateNavigation();
  const productCards = ['produtoAtualCard', 'produtoTesteCard'].map(id => root.getElementById(id)!);
  const productGrid = productCards[0].parentElement!;
  // Stable primary product first; the isolated test retains every source element.
  productGrid.append(...productCards);
  productCards.forEach((card, index) => {
    card.classList.add('product-panel');
    const top = card.querySelector<HTMLElement>('.produto-card-top')!;
    card.querySelector('.produto-card-label')!.classList.add('card-title');
    const eyebrow = document.createElement('div'); eyebrow.className = 'product-eyebrow'; eyebrow.textContent = index ? 'EXPERIMENTO' : 'PRODUTO PRINCIPAL'; card.prepend(eyebrow);
    const plot = card.querySelector<HTMLElement>('.produto-chart')!; plot.classList.add('product-plot');
    const plotHeading = document.createElement('div'); plotHeading.className = 'product-chart-header card-header'; plotHeading.innerHTML = '<span>Vendas e CAC</span>'; plot.before(plotHeading);
  });
  const testHint = document.createElement('p'); testHint.className = 'product-data-note'; testHint.textContent = 'CAC indisponível nesta seleção. Os demais dados disponíveis continuam apresentados.'; testHint.hidden = true;
  root.getElementById('produtoTesteCard')!.querySelector('.produto-card-top')!.append(testHint);
  for (const [index, card] of productCards.entries()) {
    const details = document.createElement('details'); details.className = 'product-method';
    const summary = document.createElement('summary'); summary.textContent = index ? 'Referências e atribuição do teste' : 'Como comparar este produto'; details.append(summary);
    if (index) { for (const selector of ['.cac-regua-wrap', '.produto-attrib']) { const node = card.querySelector(selector); if (node) details.append(node); } }
    else { const note = document.createElement('p'); note.textContent = 'Receita e lucro por comprador incluem os upsells. O CAC considera a consulta base, descontando o teste conforme os dados disponíveis.'; details.append(note); }
    card.querySelector('.produto-metrics')!.after(details);
  }

  const ads = root.getElementById('adsBalanceCard')!;
  const history = ads.querySelector('.ads-balance-spark');
  if (history) {
    const historyPanel = document.createElement('div'); historyPanel.className = 'balance-history';
    const caption = document.createElement('span'); caption.textContent = 'Histórico do saldo · últimas leituras'; historyPanel.append(caption, history); ads.appendChild(historyPanel);
  }
  const operationalDetails = document.createElement('details'); operationalDetails.className = 'operation-context';
  operationalDetails.innerHTML = '<summary>Entender a autonomia estimada</summary><p>Os dias restantes relacionam o saldo disponível ao ritmo recente de gasto. Uma mudança no investimento pode antecipar ou adiar o esgotamento. Confira o horário da última leitura antes de decidir.</p><a href="#/marketing">Abrir campanhas ↗</a>';
  ads.append(operationalDetails);
  const monitor = ads.parentElement!; monitor.classList.add('operation-panels');
  const bureauLabel = root.getElementById('bureauCard')!.querySelector('.bureau-type-divider > .bureau-card-label');
  text(bureauLabel, 'Composição por fornecedor');

  const profitGrid = document.createElement('div'); profitGrid.className = 'daily-profit-grid';
  const profitCards = ['profitChart', 'netProfitChart'].map(id => root.getElementById(id)!.closest<HTMLElement>('.card')!);
  profitCards[0].before(profitGrid); profitGrid.append(...profitCards);
  const switcher = document.createElement('div'); switcher.className = 'result-switcher'; switcher.setAttribute('role', 'group'); switcher.setAttribute('aria-label', 'Tipo de resultado diário');
  for (const [index, label] of ['Bruto', 'Líquido'].entries()) {
    const button = document.createElement('button'); button.textContent = label; button.setAttribute('aria-pressed', String(index === 1));
    button.onclick = () => { const header = profitCards[index].querySelector('.card-header')!; header.append(switcher); profitCards.forEach((card, i) => card.hidden = i !== index); switcher.querySelectorAll('button').forEach((b, i) => b.setAttribute('aria-pressed', String(i === index))); window.dispatchEvent(new Event('resize')); };
    switcher.append(button);
  }
  profitCards[1].querySelector('.card-header')!.append(switcher); profitCards[0].hidden = true;
  const performanceHeading = root.getElementById('section-performance')!;
  const performanceDescription = performanceHeading.nextElementSibling!;
  const financial = root.querySelector('.financial-scroll')!;
  financial.after(performanceHeading, performanceDescription, profitGrid);
  profitGrid.after(root.querySelector('.executive-analysis')!);
  root.getElementById('gaugeRow')?.after(root.querySelector('.result-goal')!);
  const charts: [string, string][] = [['profitChart', 'Total bruto no período'], ['netProfitChart', 'Resultado do painel no período'], ['cacChart', 'CAC consolidado do período'], ['upsellChart', 'Taxa consolidada do período']];
  for (const [id, caption] of charts) {
    const canvas = root.getElementById(id)!; const card = canvas.closest<HTMLElement>('.card')!;
    card.classList.add('analytical-panel');
    const main = document.createElement('div'); main.className = 'analytical-summary'; main.dataset.chart = id;
    const value = document.createElement('strong'); value.textContent = '—';
    const label = document.createElement('span'); label.textContent = caption; main.append(value, label);
    card.querySelector('.card-header')!.after(main);
    canvas.parentElement!.classList.add('analytical-plot');
    if (id === 'profitChart' || id === 'netProfitChart') {
      const legend = document.createElement('div'); legend.className = 'result-legend';
      legend.innerHTML = '<span><i class="positive-dot"></i>Positivo</span><span><i class="negative-dot"></i>Negativo</span><span><i class="mean-dot"></i>Média tracejada</span>';
      main.after(legend);
    }
  }
  const acquisitionGrid = root.getElementById('cacChart')!.closest<HTMLElement>('.grid-1-1');
  acquisitionGrid?.classList.add('acquisition-grid');

  const days = root.getElementById('dowGrid')!;
  const rhythm = days.closest<HTMLElement>('.grid-4-1')!; rhythm.classList.add('weekly-layout');
  const dayPanel = days.parentElement!; dayPanel.classList.add('weekday-panel');
  const dayHeading = document.createElement('div'); dayHeading.className = 'weekday-heading';
  dayHeading.innerHTML = '<h3>O que cada dia costuma entregar</h3><p>Resultado médio após imposto, com receita e resultado bruto para conferência.</p>';
  dayPanel.prepend(dayHeading);
  const campaignCard = root.getElementById('convChart')!.closest<HTMLElement>('.card')!; campaignCard.classList.add('campaign-distribution');
  const attribution = document.createElement('p'); attribution.className = 'attribution-note'; attribution.textContent = 'Fora do teste, pedidos são estimados pela participação nas conversões Ads. Não representa atribuição direta de cada pedido.'; campaignCard.appendChild(attribution);
  root.getElementById('convChart')!.parentElement!.classList.add('campaign-plot');


  // Group complete subjects in reading order, keeping every engine-owned node.
  const heading = (key: string) => overview.querySelector<HTMLElement>('[data-section-key="' + key + '"]')!;
  const block = (label: HTMLElement) => [label, label.nextElementSibling!];
  const efficiencyHeading = heading('Saúde'); efficiencyHeading.id = 'section-efficiency';
  const rhythmHeading = heading('Performance por Dia');
  const funnelHeading = heading('Funil');
  const refundHeading = heading('Estornos');
  const funnel = root.getElementById('funnelWrap')!.closest<HTMLElement>('.card')!;
  const conversion = document.createElement('div'); conversion.className = 'conversion-layout';
  const funnelPanel = document.createElement('div'); funnelPanel.className = 'conversion-funnel';
  funnelPanel.append(...block(funnelHeading), funnel);
  const legacyFunnel = funnel.querySelector<HTMLElement>('#funnelWrap')!;
  legacyFunnel.setAttribute('hidden', '');
  const funnelHost = document.createElement('div'); funnelHost.className = 'flow-funnel-host';
  legacyFunnel.after(funnelHost);
  const funnelRoot = createRoot(funnelHost);
  cleanups.push(() => { queueMicrotask(() => funnelRoot.unmount()); });
  let funnelSignature = '';
  conversion.append(funnelPanel, campaignCard);
  const group = (name: string, nodes: Element[]) => {
    const section = document.createElement('section'); section.className = 'overview-topic topic-' + name;
    const title = nodes[0]; if (title.id) section.setAttribute('aria-labelledby', title.id);
    section.append(...nodes); overview.append(section); return section;
  };
  group('financial', [root.getElementById('section-financial')!, financial]);
  group('performance', [...block(performanceHeading), profitGrid, root.querySelector('.executive-analysis')!, ...block(rhythmHeading), dayPanel]);
  group('efficiency', [...block(efficiencyHeading), root.querySelector('.result-goal')!, root.getElementById('gaugeRow')!, acquisitionGrid!]);
  group('products', [...block(root.getElementById('section-products')!), productGrid, conversion]);
  group('operation', [...block(root.getElementById('section-operation')!), monitor, ...block(refundHeading), root.getElementById('refundSection')!]);
  rhythm.remove();
  [rhythmHeading, funnelHeading, refundHeading].forEach(label => { label.classList.add('overview-subheading'); label.setAttribute('role','heading'); label.setAttribute('aria-level','3'); });
  root.querySelectorAll('.overview-topic > .section-label:not(.overview-subheading)').forEach(label => { label.setAttribute('role','heading'); label.setAttribute('aria-level','2'); });
  const alert = document.createElement('button'); alert.type = 'button'; alert.className = 'operation-priority'; alert.hidden = true;
  alert.onclick = () => navButtons.get('operation')!.click(); sectionNav.after(alert);

  const update = () => {
    scheduleNavigation();
    if (root.querySelectorAll('#kpiRow1 .kpi-value').length !== 5) return;
    const { totals, funnel: funnelValues } = read();
    const signature = JSON.stringify(funnelValues);
    if (signature !== funnelSignature) {
      funnelSignature = signature;
      const labels = ['Checkouts', 'Pedidos', 'Upsells'];
      const colors = ['#78c9d3', '#a4d956', '#b1a0df', '#78c9d3'];
      funnelRoot.render(createElement(FunnelChart, { data: labels.map((label, i) => ({ label, value: Math.max(0, Math.round(funnelValues[i] || 0)), color: colors[i], endColor: colors[i + 1] })) }));
    }
    const badge = root.getElementById('adsBalanceBadge')?.textContent?.trim() || '';
    const critical = /crítico|critico|esgotado/i.test(badge);
    if (alert.hidden === critical) alert.hidden = !critical;
    text(alert, 'Atenção ao saldo Ads · ' + badge + ' · Ver monitoramento ↗');
    const refunds = root.getElementById('refundSection')!;
    const compact = totals.refundCount === 0;
    if (refunds.classList.contains('refunds-compact') !== compact) refunds.classList.toggle('refunds-compact', compact);
    const values: Record<string, number | null> = { profitChart: totals.pft, netProfitChart: totals.netPftOpt, cacChart: totals.cac > 0 ? totals.cac : null, upsellChart: totals.conv > 0 ? totals.upsell : null };
    for (const [id] of charts) {
      const value = values[id], target = root.querySelector(`.analytical-summary[data-chart="${id}"] strong`)!;
      text(target, value === null ? '—' : id === 'upsellChart' ? value.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%' : currency(value));
      if (id === 'profitChart' || id === 'netProfitChart') {
        const state = resultStatus(value);
        if (target.getAttribute('data-result') !== state.tone) target.setAttribute('data-result', state.tone);
        target.setAttribute('aria-label', target.textContent + ' · ' + state.text);
      }
    }
    const marginReference = root.getElementById('gs1');
    if (marginReference) text(marginReference, (marginReference.textContent || '').replace('Meta:', 'Referência visual:'));
    const missingCac = root.getElementById('ptCac')!.textContent?.trim() === '—';
    if (testHint.hidden === missingCac) testHint.hidden = !missingCac;
    const upsellGauge = root.getElementById('gn2')?.closest('.gauge-card');
    const upsellTone = root.getElementById('gs2')?.textContent?.includes('Abaixo') ? 'negative' : 'positive';
    if (upsellGauge && upsellGauge.getAttribute('data-tone') !== upsellTone) upsellGauge.setAttribute('data-tone', upsellTone);
    // Remove pictograms, retaining every word/state produced by the original engine.
    for (const selector of ['.produto-card-label', '.ads-balance-label', '#bureauCard > .bureau-card-top .bureau-card-label', '#adsBalanceBadge', '#bureauCardBadge', '#ptFaseBadge']) {
      root.querySelectorAll(selector).forEach(element => {
        const clean = (element.textContent || '').replace(/[\p{Extended_Pictographic}\uFE0F]/gu, '').trim(); text(element, clean);
      });
    }
    const refundValues = root.querySelectorAll<HTMLElement>('#refundSection .kpi-value');
    if (refundValues.length === 3) {
      const firstColor = totals.refundCount > 0 ? 'var(--red)' : 'var(--text-primary)';
      if (refundValues[0].style.color !== firstColor) refundValues[0].style.color = firstColor;
    }
  };
  return Object.assign(update, { dispose: () => cleanups.forEach(cleanup => cleanup()) });
}

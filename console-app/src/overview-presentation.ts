import { netMarginPercent } from './domain/overview-series';
import { goalProgress, resultStatus, targetStatus } from './domain/metric-presentation';

const brl = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const copy: Record<string, string> = { r: 'Total no período', ct: 'Ads + bureau', p: 'Receita menos COGS', npo: 'Resultado do painel', eb: 'Lucro líquido / receita', ro: 'Receita / (Ads + bureau)', cpv: 'Custo total / pedidos', ca: 'Ads / pedidos', ga: 'Investimento no período', gb: 'Custo no período', m: 'Lucro bruto / receita', cv: 'Pedidos / checkouts', up: 'Pedidos com adicional / pedidos', tr: 'Pedidos pagos', tk: 'Receita / pedidos' };
const setText = (el: Element | null, text: string) => { if (el && el.textContent !== text) el.textContent = text; };

/** Reuses original elements/IDs and Chart instances; never writes to sources or changes the engine. */
export function createOverviewPresentation(root: ShadowRoot, read: () => any, summary: boolean, view: string, action: (name: 'rules' | 'export') => void) {
  let goal: HTMLElement | null = null;
  if (!summary && view === 'overview') {
    const header = root.querySelector<HTMLElement>('header')!;
    const hero = document.createElement('section'); hero.className = 'overview-hero';
    hero.innerHTML = '<div><h1>Visão Geral</h1><p>Uma leitura clara da sua operação.</p></div><div class="hero-actions"></div>';
    header.before(hero);
    const actions = hero.querySelector('.hero-actions')!;
    actions.appendChild(root.getElementById('dpWrapper')!);
    const refresh = root.getElementById('refreshBtn')!; refresh.setAttribute('aria-label', 'Atualizar dados'); refresh.innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/></svg>'; actions.appendChild(refresh);
    const download = document.createElement('button'); download.className = 'hero-export'; download.type = 'button'; download.textContent = '↓ Exportar'; download.disabled = true; download.onclick = () => action('export'); actions.appendChild(download);
    const more = document.createElement('details'); more.className = 'overview-tools';
    more.innerHTML = '<summary aria-label="Filtros e informações">···</summary><div class="overview-tools-panel"><h2>Filtros e informações</h2></div>';
    const panel = more.querySelector('.overview-tools-panel')!; panel.appendChild(header);
    const nav = document.createElement('nav'); nav.setAttribute('aria-label', 'Seções da visão geral');
    for (const [id, name] of [['financial', 'Resultados'], ['performance', 'Performance'], ['efficiency', 'Eficiência'], ['products', 'Produtos'], ['operation', 'Operação']]) {
      const button = document.createElement('button'); button.textContent = name; button.onclick = () => { more.open = false; const el = root.getElementById('section-' + id); if (el) window.scrollTo({ top: el.getBoundingClientRect().top + window.scrollY - 105, behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'instant' : 'smooth' }); }; nav.appendChild(button);
    }
    panel.appendChild(nav);
    const rules = document.createElement('button'); rules.className = 'hero-rules'; rules.textContent = 'Regras e comparativos'; rules.onclick = () => { more.open = false; action('rules'); }; panel.appendChild(rules);
    actions.appendChild(more);
    const rows = ['kpiRow1', 'kpiRow2', 'kpiRow3'].map(id => root.getElementById(id)!);
    const scroll = document.createElement('div'); scroll.className = 'financial-scroll'; scroll.tabIndex = 0;
    scroll.setAttribute('role', 'region'); scroll.setAttribute('aria-label', 'Resultados Financeiros: 15 indicadores em cinco colunas. Role para os lados em telas pequenas.');
    const grid = document.createElement('div'); grid.className = 'financial-grid';
    rows[0].before(scroll); scroll.appendChild(grid); rows.forEach(row => grid.appendChild(row));
    const lead = document.createElement('div'); lead.className = 'executive-analysis'; scroll.after(lead);
    const revenue = root.getElementById('revenueChart')?.closest<HTMLElement>('.card');
    if (revenue) {
      lead.appendChild(revenue);
      setText(revenue.querySelector('.card-title'), 'Receita e custos');
      setText(revenue.querySelector('.source-chip'), 'Fontes originais');
      const legend = document.createElement('div'); legend.className = 'revenue-legend'; legend.innerHTML = '<span><i></i>Receita</span>'; revenue.querySelector('.card-header')!.appendChild(legend);
    }
    goal = document.createElement('section'); goal.className = 'result-goal'; goal.setAttribute('aria-label', 'Progresso em relação à meta mensal de resultado');
    goal.innerHTML = '<h2>Meta de resultado</h2><strong class="goal-number">—</strong><p class="goal-amount"></p><div class="goal-track" aria-hidden="true"><span></span></div><span class="goal-status"></span><p class="goal-context"></p><a class="goal-link" href="#/metas">Conferir metas <span aria-hidden="true">↗</span></a>';
    lead.appendChild(goal);
    const scope = document.createElement('p'); scope.className = 'filter-scope';
    scope.textContent = 'Campanha: detalhamento específico. Os indicadores financeiros e gráficos mantêm o consolidado.';
    root.getElementById('section-filters')?.appendChild(scope);
  }
  return () => {
    if (root.querySelectorAll('#kpiRow1 .kpi-value').length !== 5) return;
    const snapshot = read(), t = snapshot.totals, targets = snapshot.targets;
    for (const [id, context] of Object.entries(copy)) {
      const card = root.getElementById('kc_' + id); if (!card) continue;
      if (id === 'm') setText(card.querySelector('.kpi-label'), 'Margem bruta');
      if (id === 'eb') {
        const margin = netMarginPercent(t.rev, t.netPftOpt);
        setText(card.querySelector('.kpi-label'), 'Margem líquida');
        card.querySelector('.tip-icon')?.setAttribute('aria-label', 'Informações de Margem líquida');
        (root.getElementById('kv_eb') as HTMLElement).style.removeProperty('color');
        setText(root.getElementById('kv_eb'), margin === null ? '—' : margin.toLocaleString('pt-BR', {maximumFractionDigits:1}) + '%');
        setText(root.getElementById('ksub_eb'), '');
        const formula = 'Lucro líquido do painel ÷ receita bruta. Após imposto, Ads e bureau; antes das despesas complementares da DRE.';
        card.setAttribute('data-original-tip', formula);
        root.getElementById('kv_eb')!.setAttribute('data-tone', margin === null ? 'neutral' : margin < 0 ? 'negative' : 'positive');
      }
      if (!card.querySelector('.metric-context')) {
        const label = document.createElement('div'); label.className = 'metric-context'; label.textContent = context;
        root.getElementById('kv_' + id)!.after(label);
        const states = document.createElement('div'); states.className = 'metric-states';
        label.after(states);
        const delta = root.getElementById('kd_' + id)!, badge = root.getElementById('kb_' + id)!;
        states.append(delta, badge);
        const detail = document.createElement('span'); detail.className = 'metric-detail'; detail.tabIndex = 0; card.appendChild(detail);
        if (!card.querySelector('.tip-icon')) { const info = document.createElement('button'); info.type = 'button'; info.className = 'tip-icon'; info.textContent = 'i'; info.setAttribute('aria-label', 'Informações de ' + card.querySelector('.kpi-label')!.textContent); card.classList.add('has-tip'); card.appendChild(info); }
      }
      const value = ({ r: t.rev, ct: t.costTotal, p: t.pft, npo: t.netPftOpt, ca: t.cac > 0 ? t.cac : null, up: t.upsell } as Record<string, number | null>)[id];
      const target = ({ r: targets.revenue, ct: targets.budget, p: targets.profit, npo: targets.netProfit, ca: targets.cpa, up: targets.upsell } as Record<string, number>)[id];
      let status = targetStatus(value ?? null, target ?? null, id === 'ca' || id === 'ct');
      if (['r', 'ct', 'p', 'npo'].includes(id) && snapshot.months !== 1) status = { tone: 'neutral', text: 'Meta mensal · período cruza meses' };
      else if (['r', 'p', 'npo'].includes(id) && target > 0) status.text = (value! >= target ? '✓ Meta mensal atingida' : 'Abaixo da meta mensal');
      else if (id === 'ct' && target > 0) status.text = value! <= target ? '✓ Dentro do orçamento mensal' : '! Acima do orçamento mensal';
      if (id === 'ro') status = { tone: 'neutral', text: 'Meta não definida' };
      if (id === 'eb') status = { tone: t.rev <= 0 ? 'neutral' : t.netPftOpt < 0 ? 'negative' : 'positive', text: t.rev <= 0 ? 'Sem receita no período' : t.netPftOpt < 0 ? 'Resultado negativo' : 'Resultado positivo' };
      if (id === 'up' && !t.conv) status = { tone: 'neutral', text: 'Sem pedidos no período' };
      if (id === 'm') status = { tone: t.margin >= .22 ? 'positive' : 'attention', text: t.margin >= .22 ? '✓ Referência saudável ≥22%' : '! Abaixo da referência de 22%' };
      if (['cpv', 'ga', 'gb', 'tr', 'tk'].includes(id)) status = { tone: 'neutral', text: '' };
      if (id === 'cv') status = { tone: 'neutral', text: t.convRate === null ? 'Sem dados de checkout' : '' };
      if (id === 'm' && !t.rev) status = { tone: 'neutral', text: 'Sem receita no período' };
      const badge = root.getElementById('kb_' + id)!;
      if (badge.getAttribute('data-tone') !== status.tone) badge.setAttribute('data-tone', status.tone);
      const result = ['p', 'npo'].includes(id) ? resultStatus(value ?? null) : null;
      if (result) { const el = root.getElementById('kv_' + id)!; el.setAttribute('data-tone', result.tone); el.setAttribute('aria-label', el.textContent + ' · ' + result.text); }
      const delta = root.getElementById('kd_' + id)!;
      if (['ct', 'ga', 'gb'].includes(id)) delta.setAttribute('data-cost', 'true');
      if (delta.textContent !== delta.getAttribute('data-presented')) delta.setAttribute('data-original-delta', delta.textContent || '');
      const rawDelta = delta.getAttribute('data-original-delta') || '';
      const originalSub = root.getElementById('ksub_' + id)?.textContent || '';
      const secondary = root.getElementById('kv2_' + id)?.textContent || '';
      const baseTip = card.getAttribute('data-original-tip') ?? card.getAttribute('data-tip') ?? context;
      if (!card.hasAttribute('data-original-tip')) card.setAttribute('data-original-tip', baseTip);
      const full = [baseTip, result?.text, rawDelta, status.text, originalSub, secondary].filter(Boolean).join('\n');
      card.setAttribute('data-tip', full);
      const shortDelta = rawDelta.replace(/ vs ant\./g, '').replace(/ · .*/, '');
      const cost = ['ct', 'ga', 'gb'].includes(id);
      const deltaLabel = shortDelta ? shortDelta + (cost ? '' : delta.classList.contains('pos') ? ' · melhora' : delta.classList.contains('neg') ? ' · piora' : ' · estável') : '';
      const showDelta = ['p', 'cpv', 'ca', 'cv', 'tk'].includes(id);
      setText(delta, showDelta ? deltaLabel : ''); delta.setAttribute('data-presented', delta.textContent || '');
      const statusLabel = status.text.replace('✓ Meta mensal atingida', '✓ Meta atingida').replace('Abaixo da meta mensal', '! Fora da meta').replace('✓ Dentro do orçamento mensal', '✓ No orçamento').replace('! Acima do orçamento mensal', '! Fora do orçamento').replace('✓ Referência saudável ≥22%', '✓ Dentro da referência').replace('! Abaixo da referência de 22%', '! Fora da referência').replace('Meta mensal · período cruza meses', 'Meta mensal');
      setText(badge, ['p', 'ga', 'gb', 'tr', 'tk', 'cv', 'cpv'].includes(id) ? '' : statusLabel);
      const detail = card.querySelector('.metric-detail')!;
      const meta = id === 'ca' ? 'Teto ' + brl(targets.cpa) : id === 'up' ? 'Meta ' + targets.upsell + '%' : ['r', 'p', 'npo'].includes(id) ? 'Meta ' + brl(target) : id === 'ct' ? 'Orçamento ' + brl(targets.budget) : '';
      setText(detail, meta); detail.setAttribute('title', full); detail.setAttribute('aria-label', full);
      // Uniform miniature bars use the same original observations, without fabricated trends.
      const canvas = root.getElementById('ks_' + id) as HTMLCanvasElement | null;
      if (canvas && canvas.getAttribute('aria-label') !== 'Histórico de ' + card.querySelector('.kpi-label')!.textContent) canvas.setAttribute('aria-label', 'Histórico de ' + card.querySelector('.kpi-label')!.textContent);
    }
    if (!summary && view === 'overview') {
      const order = [['r','ct','p','npo','ro'], ['cpv','ca','ga','gb','tk'], ['m','eb','cv','up','tr']];
      order.forEach((ids, i) => { const row = root.getElementById('kpiRow' + (i+1))!; ids.forEach((key, j) => { const card = root.getElementById('kc_' + key)!; if (row.children[j] !== card) row.insertBefore(card, row.children[j] || null); }); });
    }
    const download = root.querySelector<HTMLButtonElement>('.hero-export'); if (download) download.disabled = false;
    if (!goal) return;
    const progress = snapshot.months === 1 ? goalProgress(t.netPftOpt, targets.netProfit) : null;
    setText(goal.querySelector('.goal-number'), progress ? progress.percent.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%' : '—');
    setText(goal.querySelector('.goal-amount'), progress ? brl(t.netPftOpt) + ' de ' + brl(targets.netProfit) : snapshot.months !== 1 ? 'Escolha um período dentro de um único mês.' : 'Meta mensal sem valor positivo para calcular o progresso.');
    const fill = goal.querySelector<HTMLElement>('.goal-track span')!;
    const width = (progress?.track || 0).toFixed(3).replace(/\.?0+$/, '') + '%'; if (Math.abs((parseFloat(fill.style.width) || 0) - (progress?.track || 0)) > .001) fill.style.width = width;
    setText(goal.querySelector('.goal-status'), progress ? progress.remaining > 0 ? 'Faltam ' + brl(progress.remaining) : '✓ Meta atingida · ' + brl(-progress.remaining) + ' acima' : 'Sem comparação de meta');
    const tone = progress ? progress.remaining > 0 ? 'negative' : 'positive' : 'neutral';
    if (goal.querySelector('.goal-status')!.getAttribute('data-tone') !== tone) goal.querySelector('.goal-status')!.setAttribute('data-tone', tone);
    setText(goal.querySelector('.goal-context'), (snapshot.real ? 'Fontes reais' : 'EXEMPLO · fonte indisponível') + ' · Realizado no período selecionado versus meta mensal de ' + (snapshot.month || 'referência') + '. Antes das despesas complementares da DRE.');
  };
}

(() => {
  const mode = new URLSearchParams(location.search).get('mode');
  if (mode === 'summary') document.documentElement.dataset.mode = 'summary';
  document.querySelector('header').id = 'section-filters';
  const emit = payload => parent.postMessage({ type: 'vp-overview', ...payload }, location.origin);
  const setTheme = theme => {
    document.documentElement.dataset.theme = theme === 'dark' ? 'dark' : '';
    window.dispatchEvent(new Event('vp-theme-change'));
    // Presentation only: use the same theme in the shell and the embedded view.
    try { localStorage.setItem('vp_theme', theme); } catch {}
  };
  setTheme(new URLSearchParams(location.search).get('theme') || 'dark');
  document.getElementById('campaignFilter').setAttribute('aria-label', 'Campanha');
  document.getElementById('dateFrom').setAttribute('aria-label', 'Data inicial');
  document.getElementById('dateTo').setAttribute('aria-label', 'Data final');
  document.querySelectorAll('.date-preset').forEach(el => {
    el.tabIndex = 0; el.setAttribute('role', 'button');
    el.addEventListener('keydown', event => {
      if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); el.click(); }
    });
  });
  // Preserve all builders' hidden DOM dependencies, without exposing forms.
  document.querySelectorAll('.view:not(#view-overview), .sidebar').forEach(el => { el.inert = true; el.setAttribute('aria-hidden', 'true'); });
  let scheduled = false, lastHeight = 0, lastReady = false, lastExample = null, lastPeriod = '', lastSections = '';
  function chartTable(chart) {
    const title = chart.canvas.closest('.card')?.querySelector('.card-title')?.textContent.trim() || 'Valores do gráfico';
    const headers = ['Período', ...chart.data.datasets.map(series => series.label || 'Série')];
    const rows = chart.data.labels.map((label, index) => [String(label), ...chart.data.datasets.map((series, datasetIndex) => {
      const value = series.data[index];
      if (value == null) return '—';
      const callback = chart.config.options.plugins?.tooltip?.callbacks?.label;
      if (callback) {
        const text = callback({ raw: value, dataset: series, datasetIndex, dataIndex: index, label, formattedValue: String(value) });
        const formatted = String(text || '').trim(), prefix = String(series.label || '') + ':';
        return formatted.startsWith(prefix) ? formatted.slice(prefix.length).trim() : formatted;
      }
      return Number(value).toLocaleString('pt-BR', { maximumFractionDigits: 2 });
    })]);
    emit({ chartTable: { title, headers, rows } });
  }
  function chartTools() {
    if (mode === 'summary' || !window.Chart) return;
    for (const canvas of document.querySelectorAll('#view-overview .card canvas')) {
      const card = canvas.closest('.card'), header = card?.querySelector('.card-header');
      if (!header) continue;
      canvas.setAttribute('role', 'img');
      canvas.setAttribute('aria-label', header.querySelector('.card-title')?.textContent || 'Gráfico de performance');
      let button = header.querySelector('.chart-values-button');
      if (!button) {
        button = document.createElement('button'); button.className = 'chart-values-button'; button.textContent = 'Ver valores';
        button.setAttribute('aria-label', 'Ver valores de ' + canvas.getAttribute('aria-label'));
        button.onclick = () => { const chart = Chart.getChart(canvas); if (chart) chartTable(chart); };
        header.appendChild(button);
      }
      button.disabled = !Chart.getChart(canvas);
    }
  }
  const measure = () => {
    scheduled = false;
    chartTools();
    const example = !!document.getElementById('demoDataBanner');
    if (example !== lastExample) { lastExample = example; emit({ example }); }
    const period = document.getElementById('dateRangeHint').textContent.trim();
    if (period !== lastPeriod) { lastPeriod = period; emit({ period }); }
    const sectionNodes = [...document.querySelectorAll('[id^="section-"]')].filter(node => node.getBoundingClientRect().height);
    const sections = Object.fromEntries(sectionNodes.map(node => [node.id.replace('section-', ''), node.getBoundingClientRect().top + window.scrollY]));
    const serialized = JSON.stringify(sections);
    if (serialized !== lastSections) { lastSections = serialized; emit({ sections }); }
    const ready = document.querySelectorAll('.kpi-row .kpi-value').length === 15;
    if (ready !== lastReady) { lastReady = ready; emit({ ready }); }
    const height = Math.ceil(document.querySelector('.app-content').getBoundingClientRect().height + 8);
    if (height !== lastHeight) { lastHeight = height; emit({ height }); }
  };
  const schedule = () => { if (!scheduled) { scheduled = true; requestAnimationFrame(measure); } };
  new ResizeObserver(schedule).observe(document.querySelector('.app-content'));
  new MutationObserver(schedule).observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'style'] });
  window.addEventListener('resize', schedule);
  window.addEventListener('message', event => {
    if (event.origin !== location.origin || event.source !== parent || event.data?.type !== 'vp-overview-command') return;
    if (event.data.theme) setTheme(event.data.theme);
    if (event.data.section) {
      const section = document.getElementById(`section-${event.data.section}`);
      if (section) emit({ scroll: section.getBoundingClientRect().top + window.scrollY });
    }
    if (event.data.export) {
      const quote = value => '"' + String(value || '').replace(/"/g, '""') + '"';
      const rows = [...document.querySelectorAll('#kpiRow1 .kpi-card, #kpiRow2 .kpi-card, #kpiRow3 .kpi-card')].map(card => [card.querySelector('.kpi-label')?.textContent, card.querySelector('.kpi-value')?.textContent, card.querySelector('.kpi-value-secondary')?.textContent, card.querySelector('.kpi-delta')?.textContent, card.querySelector('.kpi-sub')?.textContent].map(quote).join(';'));
      const period = `${document.getElementById('dateLabel').textContent} ${document.getElementById('dateRangeHint').textContent}`;
      const source = document.getElementById('footerSource')?.textContent || document.querySelector('footer').textContent;
      emit({ csv: ['Período;' + quote(period), 'Fonte;' + quote(source), 'Indicador;Valor;Complemento;Comparativo;Contexto', ...rows].join('\n') });
    }
  });
  const revenueHeader = document.getElementById('chartSub')?.parentElement;
  if (revenueHeader) {
    const note = document.createElement('p'); note.className = 'chart-context-note';
    note.textContent = 'Investimento = Google Ads + bureau. Linhas tracejadas indicam projeções e metas.';
    revenueHeader.appendChild(note);
  }
  schedule();
})();

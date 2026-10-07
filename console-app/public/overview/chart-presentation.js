(() => {
  if (!window.Chart) return;
  const motion = matchMedia('(prefers-reduced-motion: reduce)');
  const summary = new URLSearchParams(location.search).get('mode') === 'summary';
  const meanLines = new WeakSet();
  const color = (name, fallback) => getComputedStyle(document.documentElement).getPropertyValue(name).trim() || fallback;
  function style(chart) {
    const options = chart.config.options;
    const dark = document.documentElement.dataset.theme === 'dark';
    const muted = color('--text-secondary', '#6b7280');
    if (options.animation !== false) options.animation = motion.matches || summary ? false : { duration: 420, easing: 'easeOutQuart' };
    options.resizeDelay = 120;
    if (options.plugins?.legend?.labels) {
      options.plugins.legend.labels.color = muted;
      options.plugins.legend.labels.font = { ...options.plugins.legend.labels.font, family: 'Inter, system-ui, sans-serif' };
    }
    if (options.plugins?.tooltip) Object.assign(options.plugins.tooltip, {
      backgroundColor: dark ? '#162237' : '#172438', titleColor: '#ffffff', bodyColor: '#e6edf7',
      padding: 12, cornerRadius: 8, displayColors: true, boxPadding: 5,
      titleFont: { size: 12, weight: '600' }, bodyFont: { size: 12 },
    });
    for (const scale of Object.values(options.scales || {})) {
      if (scale.ticks) {
        scale.ticks.color = muted;
        scale.ticks.font = { ...scale.ticks.font, size: Math.max(10, scale.ticks.font?.size || 11), family: 'Inter, system-ui, sans-serif' };
      }
      if (scale.grid && scale.grid.display !== false) scale.grid.color = dark ? 'rgba(152,170,198,.12)' : 'rgba(32,54,84,.08)';
      scale.border = { ...scale.border, display: false };
    }
    // Recolor only existing series styles. Values, labels, callbacks and scales remain original.
    for (const series of chart.data.datasets) {
      if (series.borderColor === 'rgba(0,0,0,.25)') meanLines.add(series);
      if (meanLines.has(series)) series.borderColor = dark ? 'rgba(183,199,222,.55)' : 'rgba(46,67,95,.4)';
    }
    const labels = options.plugins?.datalabels;
    if (labels && !labels._vpOriginalColor) {
      labels._vpOriginalColor = labels.color || '#617085';
      labels.color = context => {
        const original = labels._vpOriginalColor;
        const raw = typeof original === 'function' ? original(context) : original;
        if (!document.documentElement.dataset.theme || typeof raw !== 'string') return raw;
        if (['#15803d', '#16a34a'].includes(raw)) return color('--green', '#59cfa1');
        if (['#b91c1c', '#dc2626'].includes(raw)) return color('--red', '#ee98a5');
        if (['#1d4ed8', '#2563eb'].includes(raw)) return color('--accent', '#2874f4');
        return raw;
      };
    }
  }
  Chart.register({ id: 'verificaPresentation', beforeInit: style, beforeUpdate: style });
  Chart.defaults.font.family = 'Inter, system-ui, sans-serif';
  const repaint = () => {
    for (const chart of Object.values(Chart.instances)) {
      style(chart);
      if (chart.canvas.getBoundingClientRect().width) chart.update('none');
    }
  };
  window.addEventListener('vp-theme-change', repaint);
  motion.addEventListener('change', repaint);
})();

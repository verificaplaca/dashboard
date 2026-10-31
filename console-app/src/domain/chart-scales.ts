/** Shared visual spacing; never changes observations, units or explicit bounds. */
export function balanceChartScales(options: any, width: number) {
  for (const [id, scale] of Object.entries<any>(options.scales || {})) {
    if (scale.display === false) continue;
    const axis = scale.axis || (id.startsWith('y') ? 'y' : id.startsWith('x') ? 'x' : '');
    if (axis === 'y' && (!scale.type || scale.type === 'linear')) {
      if (scale.max == null) scale.grace = '20%';
      if (scale.min == null || scale.min === 0) scale.beginAtZero = true;
      (scale.ticks ||= {}).maxTicksLimit = 7;
    }
    // Custom cycle ticks keep their own calendar mapping.
    if (axis === 'x' && !scale.afterBuildTicks) {
      const ticks = scale.ticks ||= {};
      ticks.maxTicksLimit = width < 600 ? 4 : 6;
      ticks.maxRotation = 0;
    }
  }
}

/** Reapply on data changes, legend toggles and responsive resizes. */
export const balancedChartScales = {
  id: 'balanced-chart-scales',
  beforeUpdate(chart: { options: any; width: number }) { balanceChartScales(chart.options, chart.width); },
};
/** Same headroom for SVG histories, with zero included and missing days ignored. */
export function paddedChartRange(values: readonly (number | null | undefined)[]) {
  const finite = values.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  const low = Math.min(0, ...finite), high = Math.max(0, ...finite);
  const padding = (high - low) * .2;
  return { min: low < 0 ? low - padding : 0, max: high > 0 ? high + padding : low < 0 ? 0 : 1 };
}

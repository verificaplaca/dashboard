type Lines = readonly string[];
export interface TooltipCopy {
  title?: Lines;
  beforeBody?: Lines;
  body?: readonly { before?: Lines; lines?: Lines; after?: Lines }[];
  afterBody?: Lines;
  footer?: Lines;
}

/** Use the existing formatted callbacks, including comparisons and footnotes. */
export function chartTooltipCopy(tooltip: TooltipCopy) {
  return {
    title: (tooltip.title || []).join(' · '),
    lines: [...tooltip.beforeBody || [], ...tooltip.body?.flatMap(row => [...row.before || [], ...row.lines || [], ...row.after || []]) || [], ...tooltip.afterBody || [], ...tooltip.footer || []].filter(line => line.trim()),
  };
}

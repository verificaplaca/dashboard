export function visibleViewport(layoutHeight: number, viewport?: { height: number; offsetTop: number }) {
  const height = viewport && Number.isFinite(viewport.height) && viewport.height > 0 ? viewport.height : layoutHeight;
  const top = viewport && Number.isFinite(viewport.offsetTop) ? Math.max(0, viewport.offsetTop) : 0;
  return { height, top, inset: Math.max(0, layoutHeight - height - top) };
}

/** Smallest scroll needed to reveal a field; never force a centered position. */
export function fieldScrollDelta(field: { top: number; bottom: number }, area: { top: number; bottom: number }) {
  if (field.bottom > area.bottom) return field.bottom - area.bottom;
  if (field.top < area.top) return field.top - area.top;
  return 0;
}

/** Theme selectors must address the Shadow DOM host, including legacy syntax. */
export function scopeOverviewThemeStyles(styles: string) {
  return styles
    .replace(/:host\((\[data-theme=(?:"dark"|'dark'|dark)\])\)/g, '$1')
    .replace(/\[data-theme=(?:"dark"|'dark'|dark)\]/g, ':host([data-theme="dark"])');
}

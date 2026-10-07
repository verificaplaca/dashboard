/** One restrained entrance per card, without hiding content before observation. */
export function observeEntrances(root: HTMLElement | ShadowRoot) {
  const preference = matchMedia('(prefers-reduced-motion: reduce)');
  const selector = '.home-hero,.home-panel,.home-goal-card,.home-shortcut,.stat,.card,.summary-strip,.instance-card,.achievement-card,.kpi-card,.produto-card,.ads-balance-card,.bureau-card,.gauge-card,.dow-cell';
  const observed = new WeakSet<Element>(), animations = new Set<Animation>();
  const observer = new IntersectionObserver(entries => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      observer.unobserve(entry.target);
      if (preference.matches) continue;
      const animation = entry.target.animate([{ opacity: .45, transform: 'translateY(12px)' }, { opacity: 1, transform: 'translateY(0)' }], { duration: 420, easing: 'cubic-bezier(.2,.65,.3,1)' });
      animations.add(animation); animation.onfinish = () => animations.delete(animation);
    }
  }, { threshold: .08, rootMargin: '0px 0px -18px 0px' });
  const scan = () => root.querySelectorAll(selector).forEach(element => { if (!observed.has(element)) { observed.add(element); observer.observe(element); } });
  const mutations = new MutationObserver(scan);
  mutations.observe(root, { subtree: true, childList: true }); scan();
  const stopMotion = () => { if (preference.matches) { animations.forEach(animation => animation.cancel()); animations.clear(); } };
  preference.addEventListener('change', stopMotion);
  return () => { observer.disconnect(); mutations.disconnect(); preference.removeEventListener('change', stopMotion); animations.forEach(animation => animation.cancel()); };
}

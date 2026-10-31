import { useEffect } from 'react';
import { fieldScrollDelta, visibleViewport } from './domain/visible-viewport';
import { mobileLayoutQuery } from './lib/responsive';
import './mobile-forms.css';

function activeField(): HTMLElement | null {
  let element = document.activeElement;
  while (element?.shadowRoot?.activeElement) element = element.shadowRoot.activeElement;
  return element instanceof HTMLElement && element.matches('input:not([type=hidden]):not([type=checkbox]):not([type=radio]),textarea,select,[contenteditable=true]') ? element : null;
}

export function installMobileViewport() {
  const root = document.documentElement, viewport = window.visualViewport;
  const compact = matchMedia(mobileLayoutQuery);
  let frame = 0, reveal = false;
  const keepVisible = (height: number, top: number) => {
    const field = activeField(); if (!field) return;
    const dialogs = document.querySelectorAll<HTMLElement>('[role="dialog"]');
    const dialog = dialogs.item(dialogs.length - 1);
    if (dialog && !dialog.contains(field)) return;
    let parent = field.parentElement;
    while (parent && parent !== document.body) {
      const style = getComputedStyle(parent);
      if (/(auto|scroll)/.test(style.overflowY) && parent.scrollHeight > parent.clientHeight) {
        const rect = parent.getBoundingClientRect();
        const delta = fieldScrollDelta(field.getBoundingClientRect(), { top: Math.max(top, rect.top) + 12, bottom: Math.min(top + height, rect.bottom) - 12 });
        if (delta) parent.scrollTop += delta;
        return;
      }
      parent = parent.parentElement;
    }
    if (!dialog) {
      const delta = fieldScrollDelta(field.getBoundingClientRect(), { top: top + 72, bottom: top + height - 16 });
      if (delta) window.scrollBy({ top: delta, behavior: 'instant' });
    }
  };
  const update = () => {
    frame = 0;
    const box = visibleViewport(window.innerHeight, viewport || undefined);
    root.style.setProperty('--vp-visible-height', `${box.height}px`);
    root.style.setProperty('--vp-visible-top', `${box.top}px`);
    root.style.setProperty('--vp-keyboard-inset', compact.matches ? `${box.inset}px` : '0px');
    if (reveal && compact.matches) keepVisible(box.height, box.top);
    reveal = false;
  };
  const schedule = (showField: boolean) => {
    reveal ||= showField;
    if (!frame) frame = requestAnimationFrame(update);
  };
  const resize = () => schedule(true), pan = () => schedule(false);
  viewport?.addEventListener('resize', resize); viewport?.addEventListener('scroll', pan);
  window.addEventListener('resize', resize); document.addEventListener('focusin', resize);
  compact.addEventListener('change', resize);
  schedule(true);
  return () => {
    cancelAnimationFrame(frame);
    viewport?.removeEventListener('resize', resize); viewport?.removeEventListener('scroll', pan);
    window.removeEventListener('resize', resize); document.removeEventListener('focusin', resize);
    compact.removeEventListener('change', resize);
    for (const name of ['--vp-visible-height', '--vp-visible-top', '--vp-keyboard-inset']) root.style.removeProperty(name);
  };
}

export default function MobileViewport() { useEffect(installMobileViewport, []); return null; }

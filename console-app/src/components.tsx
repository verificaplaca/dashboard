import React, { useLayoutEffect, useRef } from 'react';
import { ArrowUpRight, ChevronRight, Search, X, Info } from 'lucide-react';
import { statusNames } from './format';

export function Badge({ status, children }: { status?: string; children?: React.ReactNode }) { return <span className={`badge ${status || ''}`}><i />{children || statusNames[status || ''] || status}</span>; }
export function Avatar({ name, size = '', src }: { name: string; size?: string; src?: string }) { return <span className={`avatar ${size}`}>{src ? <img src={src} alt={name} /> : name.split(' ').slice(0, 2).map(n => n[0]).join('').toUpperCase()}</span>; }
export function Empty({ title = 'Nenhum registro encontrado', description = 'Ajuste os filtros ou crie um novo registro.', children }: { title?: string; description?: string; children?: React.ReactNode }) { return <div className="empty"><Search size={28} /><strong>{title}</strong><p>{description}</p>{children}</div>; }
export function PageHeading({ eyebrow, title, description, children }: { eyebrow?: string; title: string; description?: string; children?: React.ReactNode }) { return <div className="page-heading"><div>{eyebrow && <div className="eyebrow">{eyebrow}</div>}<h1>{title}</h1>{description && <p>{description}</p>}</div><div className="heading-actions">{children}</div></div>; }
export function Card({ title, subtitle, action, children, className = '' }: { title?: string; subtitle?: string; action?: React.ReactNode; children: React.ReactNode; className?: string }) { return <section className={`card ${className}`}>{title && <div className="card-heading"><div><h2>{title}</h2>{subtitle && <p>{subtitle}</p>}</div>{action}</div>}{children}</section>; }
export function Stat({ label, value, hint, icon, accent, onClick }: { label: string; value: string; hint: string; icon?: React.ReactNode; accent?: boolean; onClick?: () => void }) { return <button className={`stat ${accent ? 'accent' : ''}`} onClick={onClick}><div className="stat-label"><span>{label}</span>{icon || <ArrowUpRight size={16} />}</div><strong>{value}</strong><small>{hint}</small></button>; }
export function Progress({ value, className = '' }: { value: number; className?: string }) { return <div className={`progress ${className}`}><span style={{ width: `${Math.max(0, Math.min(100, value))}%` }} /></div>; }
export function Tabs({ values, selected, onSelect }: { values: { id: string; label: string; count?: number }[]; selected: string; onSelect: (id: string) => void }) { return <div className="tabs" role="tablist">{values.map(v => <button role="tab" aria-selected={selected === v.id} className={selected === v.id ? 'selected' : ''} key={v.id} onClick={() => onSelect(v.id)}>{v.label}{v.count != null && <span>{v.count}</span>}</button>)}</div>; }
export function Toolbar({ value, onChange, placeholder = 'Buscar registros...', children }: { value: string; onChange: (value: string) => void; placeholder?: string; children?: React.ReactNode }) { return <div className="toolbar"><label className="search-field"><Search size={16} /><input value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} aria-label={placeholder} /></label><div className="toolbar-actions">{children}</div></div>; }
export function LinkButton({ children, onClick }: { children: React.ReactNode; onClick: () => void }) { return <button className="text-button" onClick={onClick}>{children}<ChevronRight size={15} /></button>; }
export function InfoNote({ children }: { children: React.ReactNode }) { return <p className="info-note"><Info size={15} />{children}</p>; }
const dialogStack: HTMLElement[] = [];
let unlockedOverflow = '';
function useDialog(close: () => void) {
  const ref = useRef<HTMLElement>(null);
  const closeRef = useRef(close);
  closeRef.current = close;
  useLayoutEffect(() => {
    const previous = document.activeElement as HTMLElement;
    const dialog = ref.current!;
    if (!dialogStack.length) unlockedOverflow = document.body.style.overflow;
    dialogStack.push(dialog);
    const focusables = () => Array.from(dialog.querySelectorAll<HTMLElement>('button:not([disabled]),input:not([disabled]):not([type="hidden"]),select:not([disabled]),textarea:not([disabled]),a[href],[tabindex="0"]')).filter(n => n.getClientRects().length && !n.closest('[inert]'));
    const listener = (e: KeyboardEvent) => {
      if (dialogStack.at(-1) !== dialog) return;
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeRef.current(); }
      if (e.key === 'Tab') {
        const nodes = focusables();
        if (!nodes.length) { e.preventDefault(); dialog.focus(); return; }
        const first=nodes[0],last=nodes[nodes.length-1];
        if (!dialog.contains(document.activeElement)) { e.preventDefault(); (e.shiftKey ? last : first).focus(); }
        else if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
        else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      }
    };
    document.addEventListener('keydown',listener); document.body.style.overflow='hidden';
    (dialog.querySelector<HTMLElement>('[autofocus],input:not([disabled]):not([type="hidden"]),textarea:not([disabled]),select:not([disabled])') || focusables()[0] || dialog).focus({preventScroll:true});
    return () => {
      document.removeEventListener('keydown',listener);
      const wasTop = dialogStack.at(-1) === dialog;
      dialogStack.splice(dialogStack.indexOf(dialog), 1);
      if (!dialogStack.length) document.body.style.overflow = unlockedOverflow;
      if (wasTop) {
        const parent = dialogStack.at(-1);
        if(previous?.isConnected && (!parent || parent.contains(previous))) previous.focus({preventScroll:true});
        else parent?.focus({preventScroll:true});
      }
    };
  }, []);
  return ref;
}
export function Dialog({ title, close, children, wide = false }: { title: string; close: () => void; children: React.ReactNode; wide?: boolean }) { const ref = useDialog(close); return <div className="modal-backdrop" onClick={e => { if(e.target === e.currentTarget) close(); }}><section ref={ref} tabIndex={-1} className={`dialog ${wide ? 'wide' : ''}`} role="dialog" aria-modal="true" aria-label={title}><div className="dialog-heading"><h2>{title}</h2><button className="icon-button" onClick={close} aria-label="Fechar"><X size={20} /></button></div><div className="dialog-content">{children}</div></section></div>; }
export function Drawer({ title, close, children }: { title: string; close: () => void; children: React.ReactNode }) { const ref = useDialog(close); return <div className="drawer-backdrop" onClick={e => { if(e.target === e.currentTarget) close(); }}><aside ref={ref} tabIndex={-1} className="drawer" role="dialog" aria-modal="true" aria-label={title}><div className="dialog-heading"><h2>{title}</h2><button className="icon-button" onClick={close} aria-label="Fechar"><X size={20} /></button></div><div className="drawer-content">{children}</div></aside></div>; }
export type Field = { name: string; label: string; type?: string; required?: boolean; value?: string | number; options?: { value: string; label: string }[]; min?: number };
export function FormDialog({ title, close, fields, onSubmit, submitLabel = 'Salvar' }: { title: string; close: () => void; fields: Field[]; onSubmit: (values: Record<string, any>) => boolean | void; submitLabel?: string }) { return <Dialog title={title} close={close}><form onSubmit={e => { e.preventDefault(); const fd = new FormData(e.currentTarget); const values: Record<string, any> = Object.fromEntries(fd); fields.filter(f => f.type === 'number').forEach(f => { values[f.name] = Number(values[f.name]); }); if (onSubmit(values) !== false) close(); }}><div className="form-fields">{fields.map(f => <label key={f.name}>{f.label}{f.options ? <select name={f.name} defaultValue={f.value} required={f.required}>{f.options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}</select> : f.type === 'textarea' ? <textarea name={f.name} defaultValue={f.value} required={f.required} rows={3} /> : <input name={f.name} type={f.type || 'text'} step={f.type === 'number' ? 'any' : undefined} defaultValue={f.value} min={f.min} required={f.required} />}</label>)}</div><div className="dialog-footer"><button type="button" className="button secondary" onClick={close}>Cancelar</button><button className="button primary" type="submit">{submitLabel}</button></div></form></Dialog>; }

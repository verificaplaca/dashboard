/** Leases prevent one completed request from hiding another request's loader. */
export function createLoadingTracker() {
 const tokens = new Set<symbol>(), listeners = new Set<() => void>();
 const publish = () => listeners.forEach(listener => listener());
 return {
  subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  getSnapshot: () => tokens.size > 0,
  begin() { const token=Symbol(); tokens.add(token); publish(); let released=false; return () => { if(released)return;released=true;tokens.delete(token);publish(); }; },
 };
}

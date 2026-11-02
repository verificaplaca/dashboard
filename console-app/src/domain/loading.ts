/** Leases prevent one completed request from hiding another request's loader. */
export function createLoadingTracker() {
 const tokens = new Set<symbol>(), listeners = new Set<() => void>();
 let total=0,completed=0;
 let progress={pending:false,total,completed};
 const publish = () => { progress={pending:tokens.size>0,total,completed};listeners.forEach(listener => listener()); };
 return {
  subscribe(listener: () => void) { listeners.add(listener); return () => { listeners.delete(listener); }; },
  getSnapshot: () => tokens.size > 0,
  getProgressSnapshot: () => progress,
  begin() { if(!tokens.size){total=0;completed=0;}total++;const token=Symbol(); tokens.add(token); publish(); let released=false; return () => { if(released)return;released=true;tokens.delete(token);completed++;publish(); }; },
 };
}

/** Queries expose completion, not transferred bytes. Keep the estimate below 100. */
export function estimateLoadingProgress(previous:number,elapsedMs:number,completed:number,total:number) {
 const elapsedEstimate=90*(1-Math.exp(-Math.max(0,elapsedMs)/6500));
 const completedEstimate=total>0?90*Math.min(1,Math.max(0,completed)/total):0;
 return Math.min(95,Math.max(previous,elapsedEstimate,completedEstimate));
}

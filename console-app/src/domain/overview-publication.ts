export const overviewUnavailableMessage = 'Não foi possível atualizar os dados. Tente novamente.';

/** An old refresh can retain real=true even after the legacy fallback runs. */
export function overviewReadState(source: {exampleBanner:boolean;ready:boolean;real:boolean;pendingReads:number}): 'loading'|'ready'|'unavailable' {
 if(source.exampleBanner)return 'unavailable';
 if(source.pendingReads>0||!source.ready)return 'loading';
 return source.real?'ready':'unavailable';
}

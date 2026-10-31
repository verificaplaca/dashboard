class PublicReadFailure extends Error {
 constructor(readonly status:number){super(`Fonte indisponível (HTTP ${status}). Tente atualizar.`);}
}
const pause=(signal?:AbortSignal)=>new Promise<void>((resolve,reject)=>{
 if(signal?.aborted){reject(signal.reason);return;}
 const stop=()=>{clearTimeout(timer);reject(signal?.reason);};
 const timer=setTimeout(()=>{signal?.removeEventListener('abort',stop);resolve();},350);
 signal?.addEventListener('abort',stop,{once:true});
});
/** Retry a transient public GET once. Authorization and malformed data never retry. */
export async function readPublicRows<T>(url:string,headers:Record<string,string>,signal?:AbortSignal,request:typeof fetch=fetch):Promise<T[]>{
 for(let attempt=0;attempt<2;attempt++){
  try{
   const deadline=AbortSignal.timeout(15_000),readSignal=signal?AbortSignal.any([signal,deadline]):deadline;
   const response=await request(url,{headers,signal:readSignal});
   if(!response.ok)throw new PublicReadFailure(response.status);
   const data:unknown=await response.json();
   if(!Array.isArray(data))throw new Error('Formato inesperado na consulta. Tente atualizar.');
   return data as T[];
  }catch(error){
   if(signal?.aborted)throw error;
   const transient=error instanceof PublicReadFailure?(error.status===429||error.status>=500):error instanceof TypeError||(error instanceof DOMException&&['TimeoutError','AbortError'].includes(error.name));
   if(attempt||!transient)throw error;
   await pause(signal);
  }
 }
 throw new Error('Não foi possível consultar a fonte.');
}

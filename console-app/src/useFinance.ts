import { useGlobalLoading } from './Loading';
import { useEffect, useRef, useState } from 'react';
import { getAuthClient } from './api';
import { useOptionalAuth } from './Auth';
import { readExpenseMonth } from './domain/finance-api';
import { loadFinanceData } from './domain/finance-data';
import type { ExpenseMonth, FinanceSnapshot } from './domain/finance';
export function useFinanceSource(start:string,end:string){
 const key=start+':'+end,[revision,setRevision]=useState(0),[state,setState]=useState<{key:string;snapshot:FinanceSnapshot|null;loading:boolean;error:string}>({key:'',snapshot:null,loading:true,error:''});
 useEffect(()=>{if(!start||!end){setState({key,snapshot:null,loading:false,error:''});return;}const controller=new AbortController();setState({key,snapshot:null,loading:true,error:''});loadFinanceData(start,end,controller.signal).then(snapshot=>{if(!controller.signal.aborted)setState({key,snapshot,loading:false,error:''});}).catch(e=>{if(!controller.signal.aborted)setState({key,snapshot:null,loading:false,error:e instanceof Error?e.message:'Falha ao carregar fontes.'});});return()=>controller.abort();},[start,end,revision]);
 const current=state.key===key?state:{snapshot:null,loading:!!start&&!!end,error:''};
 useGlobalLoading(current.loading);
 return {...current,refresh:()=>setRevision(n=>n+1)};
}
type ExpenseStatus='signed-out'|'loading'|'ready'|'forbidden'|'error';
export function useExpenseMonth(month:string){
 const auth=useOptionalAuth(),key=(auth?.user?.id??'anon')+':'+month,[state,setState]=useState<{key:string;row:ExpenseMonth|null;status:ExpenseStatus;error:string}>({key:'',row:null,status:'signed-out',error:''});const revision=useRef(0);
 const reload=async()=>{const id=++revision.current;if(!auth?.user){setState({key,row:null,status:'signed-out',error:''});return null;}setState({key,row:null,status:'loading',error:''});try{const result=await readExpenseMonth(await getAuthClient(),month);if(id===revision.current)setState({key,row:result.row,status:result.allowed?'ready':'forbidden',error:''});return result;}catch(e){if(id===revision.current)setState({key,row:null,status:'error',error:e instanceof Error?e.message:'Falha ao ler despesas.'});return null;}};
 useEffect(()=>{void reload();return()=>{revision.current++;};},[month,auth?.user?.id]);
 const current=state.key===key?state:{row:null,status:auth?.user?'loading' as const:'signed-out' as const,error:''};
 useGlobalLoading(current.status==='loading');
 return {...current,reload,auth};
}

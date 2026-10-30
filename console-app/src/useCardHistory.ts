import { useGlobalLoading } from './Loading';
import { useEffect, useState } from 'react';
import { loadCardHistory, type CardHistory } from './domain/card-history';
export function useCardHistory(start:string,end:string,revision:string){
 const key=start+':'+end+':'+revision,[state,setState]=useState<{key:string;data:CardHistory|null;loading:boolean}>({key:'',data:null,loading:true});
 useEffect(()=>{if(!revision){setState({key,data:null,loading:false});return;}const controller=new AbortController();setState({key,data:null,loading:true});loadCardHistory(start,end,controller.signal).then(data=>{if(!controller.signal.aborted)setState({key,data,loading:false});}).catch(()=>{if(!controller.signal.aborted)setState({key,data:null,loading:false});});return()=>controller.abort();},[key]);
 const current=state.key===key?state:{key,data:null,loading:!!revision};
 useGlobalLoading(current.loading);
 return current;
}

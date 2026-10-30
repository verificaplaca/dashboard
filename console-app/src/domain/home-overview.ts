export type HomeTargets={month:string;revenue_target:number|null;target_cpa:number|null;monthly_budget:number|null};
export function homeTarget(rows:HomeTargets[]|null,month:string,key:Exclude<keyof HomeTargets,'month'>){
 const row=rows?.filter(row=>row.month<=month+'-01'&&row[key]!==null).sort((a,b)=>a.month.localeCompare(b.month)).at(-1);
 const value=row?.[key];return typeof value==='number'&&Number.isFinite(value)&&value>=0?value:null;
}
export type HomeDispatch={checkout_id:string;ga4_status:string;ads_status:string;meta_status:string;created_at:string};
export function dispatchSummary(rows:HomeDispatch[]){
 let failed=0,pending=0,unknown=0;
 for(const row of rows){const statuses=[row.ga4_status,row.ads_status,row.meta_status];
  if(statuses.includes('failed'))failed++;
  else if(statuses.includes('pending'))pending++;
  else if(statuses.some(status=>!['success','skipped'].includes(status)))unknown++;
 }
 return {total:rows.length,failed,pending,unknown};
}

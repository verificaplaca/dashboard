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

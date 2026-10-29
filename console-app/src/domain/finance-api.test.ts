import test from 'node:test';
import assert from 'node:assert/strict';
import { FinanceConflict, readExpenseMonth, saveExpenseMonth } from './finance-api';
const values={fees:0,people:0,systems:null,other:0,financial:0};
test('private expenses are not read when access is denied and backend failures are never treated as zero',async()=>{
 let read=false;const denied:any={rpc:async()=>({data:false,error:null}),from:()=>{read=true;throw Error('Unexpected');}};
 assert.deepEqual(await readExpenseMonth(denied,'2026-10'),{allowed:false,row:null});assert.equal(read,false);
 await assert.rejects(readExpenseMonth({rpc:async()=>({error:{message:'offline'}})} as any,'2026-10'));
 const broken:any={rpc:async()=>({data:true}),from:()=>({select:()=>({eq:()=>({maybeSingle:async()=>({error:{message:'offline'}})})})})};await assert.rejects(readExpenseMonth(broken,'2026-10'));
});
test('saving uses only the dedicated RPC with expected version, never silently retries conflicts or grants access',async()=>{
 let calls=0;const client:any={rpc:async(name:string,args:any)=>{calls++;assert.equal(name,'console_save_finance_month');assert.deepEqual(args,{p_month:'2026-10-01',p_values:values,p_version:2});return {error:{code:'40001'}};}};
 await assert.rejects(saveExpenseMonth(client,'2026-10',values,2),FinanceConflict);assert.equal(calls,1);
 await assert.rejects(saveExpenseMonth({rpc:async()=>({error:{code:'42501'}})} as any,'2026-10',values,2),/permissão/);
});

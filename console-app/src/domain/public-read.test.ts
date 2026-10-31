import test from 'node:test';
import assert from 'node:assert/strict';
import {readPublicRows} from './public-read';
test('a transient service failure retries once and preserves the returned records',async()=>{
 let calls=0;const rows=[{id:'real-order',value:14.99}];
 const request=(async()=>++calls===1?new Response('',{status:503}):Response.json(rows)) as typeof fetch;
 assert.deepEqual(await readPublicRows('/read',{},undefined,request),rows);assert.equal(calls,2);
});
test('authorization failures and invalid payloads are not retried',async()=>{
 for(const response of [new Response('',{status:403}),Response.json({error:'invalid'})]){
  let calls=0;await assert.rejects(readPublicRows('/read',{},undefined,(async()=>{calls++;return response;}) as typeof fetch));assert.equal(calls,1);
 }
});
test('cancelling during the retry delay stops without another request',async()=>{
 const controller=new AbortController();let calls=0;
 const promise=readPublicRows('/read',{},controller.signal,(async()=>{calls++;setTimeout(()=>controller.abort(),10);return new Response('',{status:503});}) as typeof fetch);
 await assert.rejects(promise);assert.equal(calls,1);
});
test('repeated transient failure terminates without inventing rows',async()=>{
 let calls=0;await assert.rejects(readPublicRows('/read',{},undefined,(async()=>{calls++;return new Response('',{status:500});}) as typeof fetch),/HTTP 500/);assert.equal(calls,2);
});

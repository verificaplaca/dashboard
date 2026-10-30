import test from 'node:test';
import assert from 'node:assert/strict';
import {dispatchSummary,type HomeDispatch} from './home-overview';
const row=(statuses:string[])=>({ga4_status:statuses[0],ads_status:statuses[1],meta_status:statuses[2]}) as HomeDispatch;
test('dispatch counts are by event, prioritize failures and preserve unknown/skipped states',()=>{
 assert.deepEqual(dispatchSummary([row(['failed','failed','pending']),row(['success','pending','skipped']),row(['skipped','skipped','skipped']),row(['success','new-status','success'])]),{total:4,failed:1,pending:1,unknown:1});
 assert.deepEqual(dispatchSummary([]),{total:0,failed:0,pending:0,unknown:0});
});

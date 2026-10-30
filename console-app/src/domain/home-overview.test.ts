import test from 'node:test';
import assert from 'node:assert/strict';
import {homeTarget,dispatchSummary,type HomeTargets,type HomeDispatch} from './home-overview';
const targets:HomeTargets[]=[{month:'2026-09-01',revenue_target:1000,target_cpa:9,monthly_budget:200},{month:'2026-10-01',revenue_target:null,target_cpa:0,monthly_budget:300},{month:'2026-11-01',revenue_target:2000,target_cpa:12,monthly_budget:400}];
test('home goals inherit per field and never use a future month or replace an explicit zero',()=>{
 assert.equal(homeTarget(targets,'2026-10','revenue_target'),1000);assert.equal(homeTarget(targets,'2026-10','target_cpa'),0);assert.equal(homeTarget(targets,'2026-10','monthly_budget'),300);assert.equal(homeTarget(targets,'2026-08','revenue_target'),null);assert.equal(homeTarget(null,'2026-10','revenue_target'),null);
});
test('invalid goal values remain unavailable',()=>{for(const value of [NaN,Infinity,-1])assert.equal(homeTarget([{...targets[0],revenue_target:value}],'2026-10','revenue_target'),null);});
const row=(statuses:string[])=>({ga4_status:statuses[0],ads_status:statuses[1],meta_status:statuses[2]}) as HomeDispatch;
test('dispatch counts are by event, prioritize failures and preserve unknown/skipped states',()=>{
 assert.deepEqual(dispatchSummary([row(['failed','failed','pending']),row(['success','pending','skipped']),row(['skipped','skipped','skipped']),row(['success','new-status','success'])]),{total:4,failed:1,pending:1,unknown:1});
 assert.deepEqual(dispatchSummary([]),{total:0,failed:0,pending:0,unknown:0});
});

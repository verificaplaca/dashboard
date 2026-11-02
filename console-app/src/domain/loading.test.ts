import test from 'node:test';
import assert from 'node:assert/strict';
import {createLoadingTracker,estimateLoadingProgress} from './loading';
test('the overlay waits for all concurrent screen reads',()=>{
 const tracker=createLoadingTracker(),first=tracker.begin(),second=tracker.begin();
 first();assert.equal(tracker.getSnapshot(),true);
 second();assert.equal(tracker.getSnapshot(),false);
});
test('unmount or cancellation releases only its own lease and cleanup is idempotent',()=>{
 const tracker=createLoadingTracker(),oldRoute=tracker.begin(),newRoute=tracker.begin();
 oldRoute();oldRoute();assert.equal(tracker.getSnapshot(),true);
 newRoute();assert.equal(tracker.getSnapshot(),false);
});
test('a failed read can release in finally and a later retry works',async()=>{
 const tracker=createLoadingTracker(),finish=tracker.begin();
 await assert.rejects((async()=>{try{throw new Error('offline');}finally{finish();}})(),/offline/);
 assert.equal(tracker.getSnapshot(),false);
 const retry=tracker.begin();assert.equal(tracker.getSnapshot(),true);retry();assert.equal(tracker.getSnapshot(),false);
});
test('unsubscribed screens are not notified and a new subscriber sees current work',()=>{
 const tracker=createLoadingTracker();let notifications=0;
 const unsubscribe=tracker.subscribe(()=>notifications++),release=tracker.begin();
 assert.equal(notifications,1);unsubscribe();release();assert.equal(notifications,1);
 const active=tracker.begin();assert.equal(tracker.getSnapshot(),true);active();
});
test('progress snapshots are stable and count each completed lease once',()=>{
 const tracker=createLoadingTracker();assert.equal(tracker.getProgressSnapshot(),tracker.getProgressSnapshot());
 const first=tracker.begin(),second=tracker.begin();
 assert.deepEqual(tracker.getProgressSnapshot(),{pending:true,total:2,completed:0});
 first();first();assert.deepEqual(tracker.getProgressSnapshot(),{pending:true,total:2,completed:1});
 second();assert.deepEqual(tracker.getProgressSnapshot(),{pending:false,total:2,completed:2});
 const next=tracker.begin();assert.deepEqual(tracker.getProgressSnapshot(),{pending:true,total:1,completed:0});next();
});
test('estimated progress advances during a slow request but cannot claim completion',()=>{
 assert.equal(estimateLoadingProgress(0,0,0,1),0);
 assert.ok(estimateLoadingProgress(0,5000,0,1)>0);
 assert.ok(estimateLoadingProgress(0,3600000,0,1)<100);
 assert.ok(estimateLoadingProgress(0,3600000,1,1)<100);
});
test('finished work advances progress and discovering more work never moves it backwards',()=>{
 const first=estimateLoadingProgress(0,100,1,2);
 assert.equal(first,45);
 assert.equal(estimateLoadingProgress(first,200,1,4),first);
 assert.ok(estimateLoadingProgress(first,500,3,4)>first);
});
test('an idle gap or a chained request keeps the visible estimate monotonic',()=>{
 const tracker=createLoadingTracker(),finish=tracker.begin();finish();
 const before=estimateLoadingProgress(60,7000,1,1),next=tracker.begin(),state=tracker.getProgressSnapshot();
 assert.equal(estimateLoadingProgress(before,7100,state.completed,state.total),before);next();
});

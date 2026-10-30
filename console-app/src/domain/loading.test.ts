import test from 'node:test';
import assert from 'node:assert/strict';
import {createLoadingTracker} from './loading';
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

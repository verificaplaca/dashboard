import test from 'node:test';
import assert from 'node:assert/strict';
import {overviewReadState} from './overview-publication';
const real={exampleBanner:false,ready:true,real:true,pendingReads:0};
test('a successful real read can be published, including a genuinely empty period',()=>{
 assert.equal(overviewReadState(real),'ready');
});
test('a fallback cannot be published even when a failed refresh retains real=true',()=>{
 assert.equal(overviewReadState({...real,exampleBanner:true}),'unavailable');
 assert.equal(overviewReadState({...real,exampleBanner:true,pendingReads:2}),'unavailable');
});
test('a completed synthetic read without a banner is also rejected',()=>{
 assert.equal(overviewReadState({...real,real:false}),'unavailable');
});
test('incomplete and pending reads never publish intermediate values',()=>{
 assert.equal(overviewReadState({...real,ready:false,real:false,pendingReads:1}),'loading');
 assert.equal(overviewReadState({...real,pendingReads:1}),'loading');
});
test('a clean successful retry is accepted after a failed attempt',()=>{
 assert.equal(overviewReadState({...real,exampleBanner:true}),'unavailable');
 assert.equal(overviewReadState({...real,ready:false,pendingReads:1}),'loading');
 assert.equal(overviewReadState(real),'ready');
});

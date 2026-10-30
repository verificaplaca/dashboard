import test from 'node:test';
import assert from 'node:assert/strict';
import {createScrollLock} from '../lib/scroll-lock';
test('closing a dialog behind a pending loader does not unlock or permanently lock the page',()=>{
 const style={overflow:'auto'},lock=createScrollLock(()=>style),dialog=lock(),loader=lock();
 dialog();assert.equal(style.overflow,'hidden');loader();assert.equal(style.overflow,'auto');loader();assert.equal(style.overflow,'auto');
});
test('closing the loader first preserves the dialog lock until the dialog closes',()=>{
 const style={overflow:''},lock=createScrollLock(()=>style),dialog=lock(),loader=lock();
 loader();assert.equal(style.overflow,'hidden');dialog();assert.equal(style.overflow,'');
});

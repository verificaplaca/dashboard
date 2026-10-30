import test from 'node:test';
import assert from 'node:assert/strict';
import {readOverviewPeriod,validateOverviewPeriod} from './overview-period';
test('shared period restores native Overview storage and safely ignores corrupt values',()=>{
 const saved={preset:'custom' as const,label:'25/10 – 30/10',from:'2026-10-25',to:'2026-10-30'};
 assert.deepEqual(readOverviewPeriod({getItem:()=>JSON.stringify(saved)}),saved);
 for(const raw of ['invalid','null','{"preset":"forged"}','{"preset":"custom","from":"2026-02-30","to":"2026-03-01"}'])assert.deepEqual(readOverviewPeriod({getItem:()=>raw}),{preset:'month'});
});
test('custom ranges reject impossible and reversed dates while allowing one day and empty-data future ranges',()=>{
 for(const [from,to] of [['2026-10-30','2026-10-29'],['2026-02-30','2026-03-01'],['','2026-10-30'],['2026-10-30','invalid']])assert.throws(()=>validateOverviewPeriod({preset:'custom',from,to}));
 for(const date of ['2026-10-30','2099-01-01'])assert.doesNotThrow(()=>validateOverviewPeriod({preset:'custom',from:date,to:date}));
});

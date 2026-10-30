import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {homeOverviewValues,type OverviewHomeSnapshot} from './overview-home-snapshot';
test('home copies the frozen Overview totals, including today, without DRE deductions',()=>{
 const html=fs.readFileSync(new URL('../../reference/dashboard-original.html',import.meta.url),'utf8');
 const calculation=html.slice(html.indexOf('function totals() {'),html.indexOf('function campTotals() {'));
 const days=[{dateStr:'2026-10-29',revenue:1000,cost:200,costBureau:100,costTotal:300,conv:50,profit:700},{dateStr:'2026-10-30',revenue:2000,cost:300,costBureau:150,costTotal:450,conv:100,profit:1550}];
 const context={filtered:days,activeCamp:false,usingRealData:true,netFactor:()=>.92,result:null as Record<string,number>|null};
 vm.runInNewContext(calculation+';result=totals();',context);
 const raw=context.result!;
 const snapshot:OverviewHomeSnapshot={real:true,totals:raw as unknown as OverviewHomeSnapshot['totals'],targets:{revenue:5000,cpa:9,budget:1000},days:days.map(day=>({date:day.dateStr,revenue:day.revenue}))};
 assert.deepEqual(homeOverviewValues(snapshot),{sales:3000,profit:2010,orders:150,media:500,cac:3.33});
 assert.equal(snapshot.days.at(-1)?.date,'2026-10-30');
 for(const [home,overview] of [['sales','rev'],['profit','netPftOpt'],['orders','conv'],['media','cost'],['cac','cac']] as const)assert.equal(homeOverviewValues(snapshot)![home],raw[overview]);
});
test('missing or example snapshots cannot appear as real home figures',()=>{
 assert.equal(homeOverviewValues(null),null);
 assert.equal(homeOverviewValues({real:false} as OverviewHomeSnapshot),null);
});
test('an empty real period preserves Overview zero values including CAC',()=>{
 const snapshot={real:true,totals:{rev:0,netPftOpt:0,conv:0,cac:0,cost:0}} as OverviewHomeSnapshot;
 assert.deepEqual(homeOverviewValues(snapshot),{sales:0,profit:0,orders:0,media:0,cac:0});
});

import { loadSimulationData, rows } from './simulation-data';
import { dates, round, taxAt, type Baseline, type Projection, type Snapshot } from './simulation';
export type UpsellDay={date:string;orders:number;revenue:number};
export type CheckoutDay={date:string;count:number};
export type CardHistory={upsell:UpsellDay[]|null;checkouts:CheckoutDay[]|null;global:{value:number;start:string;end:string;daily:number[]}|null;warnings:string[]};
const number=(v:unknown)=>{if(v==null||v===''||!Number.isFinite(Number(v))||Number(v)<0)throw new Error('Agregado inválido.');return Number(v);};
function unique<T extends {date:string}>(items:T[]){if(new Set(items.map(r=>r.date)).size!==items.length)throw new Error('Dias duplicados.');for(const r of items)dates(r.date,r.date);return items;}
export async function loadCardHistory(start:string,end:string,signal?:AbortSignal):Promise<CardHistory>{
 dates(start,end);const filter=`date=gte.${start}&date=lte.${end}&order=date.asc`;
 const results=await Promise.allSettled([
  rows('upsell_daily','date,upsell_orders,upsell_revenue',filter,signal).then(items=>unique(items.map(r=>{const orders=number(r.upsell_orders);if(!Number.isInteger(orders))throw new Error('Pedidos inválidos.');return {date:String(r.date).slice(0,10),orders,revenue:number(r.upsell_revenue)};}))),
  rows('checkouts_daily','date,checkouts_started',filter,signal,'site').then(items=>unique(items.map(r=>{const count=number(r.checkouts_started);if(!Number.isInteger(count))throw new Error('Checkouts inválidos.');return {date:String(r.date).slice(0,10),count};}))),
  loadSimulationData('2026-01-01',end,signal).then(s=>{
   if(!s.revenue.length||!s.bureau.length||!s.ads.length)throw new Error('Histórico global incompleto.');
   if(new Set(s.revenue.map(r=>r.date)).size!==s.revenue.length||new Set(s.bureau.map(r=>r.date)).size!==s.bureau.length||new Set(s.ads.map(r=>r.date+':'+r.campaign_id)).size!==s.ads.length)throw new Error('Duplicação no acumulado.');
   const days=[...new Set([...s.revenue,...s.bureau,...s.ads].map(r=>r.date))].sort();
   const daily=days.map(date=>round(s.revenue.filter(r=>r.date===date).reduce((a,r)=>a+r.revenue,0)-s.bureau.filter(r=>r.date===date).reduce((a,r)=>a+r.custo_bureau,0)-s.ads.filter(r=>r.date===date).reduce((a,r)=>a+r.cost_micros/1e6,0)));
   return {value:round(daily.reduce((a,n)=>a+n,0)),start:'2026-01-01',end,daily};
  })
 ]);
 if(signal?.aborted)throw new DOMException('Aborted','AbortError');
 const warnings=results.flatMap((r,i)=>r.status==='rejected'?[['Upsell indisponível.','Checkouts indisponíveis.','Acumulado histórico indisponível.'][i]]:[]);
 return {upsell:results[0].status==='fulfilled'?results[0].value:null,checkouts:results[1].status==='fulfilled'?results[1].value:null,global:results[2].status==='fulfilled'?results[2].value:null,warnings};
}
export function cardRates(s:Snapshot,base:Baseline,h:CardHistory|null,start:string,end:string){
 const inside=(r:{date:string})=>r.date>=start&&r.date<=end;
 const ups=h?.upsell?.filter(inside),check=h?.checkouts?.filter(inside);
 const salesDates=s.revenue.filter(r=>inside(r)&&r.paid_orders>0).map(r=>r.date);
 const upsellKnown=!!ups&&salesDates.every(d=>ups.some(r=>r.date===d));
 const upsellOrders=ups?.reduce((a,r)=>a+r.orders,0)??0,upsellRevenue=ups?.reduce((a,r)=>a+r.revenue,0)??0;
 const validUpsell=upsellKnown&&upsellOrders<=base.orders&&upsellRevenue<=base.revenue+.01;
 const checkoutKnown=!!check&&dates(start,end).every(d=>check.some(r=>r.date===d));
 const checkouts=checkoutKnown?check!.reduce((a,r)=>a+r.count,0):null;
 return {conversionPct:checkouts!==null&&checkouts>0&&base.orders<=checkouts?base.orders/checkouts*100:null,upsellPct:validUpsell?upsellOrders/base.orders*100:null,upsellUnitRevenue:validUpsell&&upsellOrders>0?upsellRevenue/upsellOrders:null,checkouts,checkoutPerReal:checkouts!==null&&base.ads>0?checkouts/base.ads:null};
}
export type CardRates=ReturnType<typeof cardRates>;
export type CardEdits={ticket?:number;bureauUnit?:number;orderRate?:number;volumeSource?:'cac'|'transactions'|'conversion';upsell?:{rate:number;unitRevenue:number;unitBureau:number}};
export function resolveCardEdits(base:Baseline,rates:CardRates,edits:CardEdits){
 if(edits.ticket!==undefined&&(!Number.isFinite(edits.ticket)||edits.ticket<.01)||edits.bureauUnit!==undefined&&(!Number.isFinite(edits.bureauUnit)||edits.bureauUnit<0))throw new Error('A alteração deixaria a base de ticket ou bureau negativa. Reveja primeiro o mix de upsell.');
 if(edits.upsell&&rates.upsellPct!==null&&((rates.upsellPct/100*edits.upsell.unitRevenue)>(edits.ticket??base.ticket)+.01||(rates.upsellPct/100*edits.upsell.unitBureau)>(edits.bureauUnit??base.bureauUnit)+.01))throw new Error('O valor incremental é incompatível com o ticket ou bureau total da base. Reveja a hipótese de upsell.');
 const change=edits.upsell?(edits.upsell.rate-(rates.upsellPct??NaN))/100:0;
 const ticket=(edits.ticket??base.ticket)+change*(edits.upsell?.unitRevenue??0),bureauUnit=(edits.bureauUnit??base.bureauUnit)+change*(edits.upsell?.unitBureau??0);
 if(!Number.isFinite(ticket)||ticket<.01||!Number.isFinite(bureauUnit)||bureauUnit<0)throw new Error('O mix de upsell informado gera ticket ou custo inválido. Reveja os valores incrementais.');
 return {ticket:edits.ticket!==undefined||edits.upsell?ticket:undefined,bureauUnit:edits.bureauUnit!==undefined||edits.upsell?bureauUnit:undefined,orderRate:edits.orderRate};
}
export type EditableCard='r'|'ct'|'ca'|'ga'|'gb'|'cv'|'up'|'tr'|'tk';
export function applyCardEdit(id:EditableCard,value:number,mode:string,base:Baseline,rates:CardRates,edits:CardEdits,p:Projection,days:number,options?:{unitRevenue:number;unitBureau:number}):{edits:CardEdits;budget?:number}{
 if(!Number.isFinite(value)||value<0||value>1e9)throw new Error('Informe um número válido, positivo ou zero.');
 const next={...edits},sales=p.orders,spend=p.ads;
 const setVolume=(orders:number)=>{if(spend<=0)throw new Error('Defina investimento positivo antes de alterar o volume.');next.orderRate=orders/spend;next.volumeSource=id==='cv'?'conversion':id==='ca'?'cac':'transactions';};
 switch(id){
  case 'ga':{const budget=value/days;if(budget>10000)throw new Error('O limite da simulação é R$ 10.000/dia.');return {edits:next,budget:round(budget)};}
  case 'ca':if(value<=0)throw new Error('CAC deve ser maior que zero.');next.orderRate=1/value;next.volumeSource='cac';break;
  case 'tr':setVolume(value);break;
  case 'cv':if(value>100||rates.checkoutPerReal===null||rates.checkoutPerReal<=0)throw new Error('Use 0–100% e uma base completa de checkouts.');setVolume(spend*rates.checkoutPerReal*value/100);break;
  case 'r':if(mode==='volume'){const ticket=sales>0?p.revenue/sales:base.ticket;setVolume(value/ticket);}else{if(sales<=0||value<=0)throw new Error('É preciso ter vendas e receita positiva para alterar o ticket.');next.ticket=value/sales-((edits.upsell?.rate??0)-(rates.upsellPct??0))/100*(edits.upsell?.unitRevenue??0);}break;
  case 'tk':if(value<.01)throw new Error('Ticket deve ser positivo.');next.ticket=value-((edits.upsell?.rate??0)-(rates.upsellPct??0))/100*(edits.upsell?.unitRevenue??0);break;
  case 'gb':case 'ct':{const bureau=id==='ct'?value-spend:value;if(bureau<0)throw new Error('O COGS não pode ficar abaixo do Ads mantido. Edite Gasto Google Ads para reduzir a mídia.');if(sales<=0)throw new Error('É preciso ter vendas para definir o custo de bureau por pedido.');next.bureauUnit=bureau/sales-((edits.upsell?.rate??0)-(rates.upsellPct??0))/100*(edits.upsell?.unitBureau??0);break;}
  case 'up':if(value>100||rates.upsellPct===null||!options||!Number.isFinite(options.unitRevenue)||options.unitRevenue<0||!Number.isFinite(options.unitBureau)||options.unitBureau<0)throw new Error('Informe a receita e o custo adicionais por pedido com upsell. Custo desconhecido não pode virar zero.');next.upsell={rate:value,...options};break;
 }
 const resolved=resolveCardEdits(base,rates,next);
 if(resolved.ticket!==undefined&&resolved.ticket>1e6||resolved.bureauUnit!==undefined&&resolved.bureauUnit>1e6||resolved.orderRate!==undefined&&resolved.orderRate>10000)throw new Error('Hipótese fora da faixa permitida.');
 return {edits:next};
}
export function financialCardValues(p:Projection,rates:CardRates,edits:CardEdits,global:CardHistory['global']){
 const cogs=round(p.ads+p.bureau),gross=round(p.revenue-cogs);
 const keepsConversion=edits.orderRate!==undefined&&edits.volumeSource!=='conversion';
 const checkouts=keepsConversion?(rates.conversionPct!==null&&rates.conversionPct>0?p.orders/(rates.conversionPct/100):null):rates.checkoutPerReal===null?null:p.ads*rates.checkoutPerReal;
 const conversion=checkouts!==null&&checkouts>0?p.orders/checkouts*100:null;
 return {r:p.revenue,ct:cogs,p:gross,npo:p.profit,eb:global?.value??null,ro:cogs>0?p.revenue/cogs:null,cpv:p.orders>0?cogs/p.orders:null,ca:p.orders>0?p.ads/p.orders:null,ga:p.ads,gb:p.bureau,ml:p.revenue>0?p.profit/p.revenue*100:null,m:p.revenue>0?gross/p.revenue*100:null,cv:conversion!==null&&conversion<=100.000001?Math.min(100,conversion):null,up:edits.upsell?.rate??rates.upsellPct,tr:p.orders,tk:p.orders>0?p.revenue/p.orders:null};
}
export type CardId=keyof ReturnType<typeof financialCardValues>;
export function cardSparks(s:Snapshot,h:CardHistory|null,start:string,end:string):Record<CardId,(number|null)[]>{
 const out=Object.fromEntries(['r','ct','p','npo','eb','ro','cpv','ca','ga','gb','m','ml','cv','up','tr','tk'].map(id=>[id,[]])) as unknown as Record<CardId,(number|null)[]>;
 for(const day of dates(start,end)){
  const revenue=s.revenue.find(r=>r.date===day),bureau=s.bureau.find(r=>r.date===day),ads=s.ads.filter(r=>r.date===day),checkout=h?.checkouts?.find(r=>r.date===day),upsell=h?.upsell?.find(r=>r.date===day);
  if(!revenue||!bureau||!ads.length){for(const id of Object.keys(out) as CardId[])out[id].push(null);continue;}
  const a=ads.reduce((n,r)=>n+r.cost_micros/1e6,0),gross=revenue.revenue-a-bureau.custo_bureau;
  const tax=taxAt(s.targets,day).value;
  const values={r:revenue.revenue,ct:a+bureau.custo_bureau,p:gross,npo:gross-revenue.revenue*tax/100,eb:null,ro:a+bureau.custo_bureau>0?revenue.revenue/(a+bureau.custo_bureau):null,cpv:revenue.paid_orders>0?(a+bureau.custo_bureau)/revenue.paid_orders:null,ca:revenue.paid_orders>0?a/revenue.paid_orders:null,ga:a,gb:bureau.custo_bureau,ml:revenue.revenue>0?(gross-revenue.revenue*tax/100)/revenue.revenue*100:null,m:revenue.revenue>0?gross/revenue.revenue*100:null,cv:checkout&&checkout.count>0?revenue.paid_orders/checkout.count*100:null,up:upsell&&revenue.paid_orders>0?upsell.orders/revenue.paid_orders*100:null,tr:revenue.paid_orders,tk:revenue.paid_orders>0?revenue.revenue/revenue.paid_orders:null};
  for(const id of Object.keys(out) as CardId[])out[id].push(values[id]);
 }
 out.eb=h?.global?.daily??[];return out;
}

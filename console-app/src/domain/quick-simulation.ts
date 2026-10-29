import { addDays, project, round, type Baseline, type TaxRow } from './simulation';
export type QuickSettings={start:string;days:number;reference:number;learning:boolean;feeRate:number;fixedDay:number;ticket?:number;bureauUnit?:number;orderRate?:number};
export const quickResponses=[{id:'careful',label:'Vendas crescem menos',elasticity:.6,color:'#a1aec3'},{id:'central',label:'Crescimento moderado',elasticity:.8,color:'#4c83ff'},{id:'same',label:'Mesma eficiência',elasticity:1,color:'#2eb797'}] as const;
export function quickProjection(base:Baseline,targets:TaxRow[],settings:QuickSettings,budget:number,elasticity:number){
 if(!Number.isFinite(budget)||budget<0||budget>10000||!Number.isInteger(settings.days)||settings.days<1||settings.days>366||!Number.isFinite(settings.reference)||settings.reference<=0||settings.reference>10000||!Number.isFinite(settings.feeRate)||settings.feeRate<0||!Number.isFinite(settings.fixedDay)||settings.fixedDay<0)throw new Error('Confira investimento e período.');
 if(settings.ticket!==undefined&&(!Number.isFinite(settings.ticket)||settings.ticket<.01||settings.ticket>1e6)||settings.bureauUnit!==undefined&&(!Number.isFinite(settings.bureauUnit)||settings.bureauUnit<0||settings.bureauUnit>1e6)||settings.orderRate!==undefined&&(!Number.isFinite(settings.orderRate)||settings.orderRate<0||settings.orderRate>10000))throw new Error('Confira as hipóteses dos cards.');
 const learns=settings.learning&&settings.orderRate===undefined&&Math.abs(budget-settings.reference)>.005,learningDays=Math.min(7,settings.days);
 const cycles=learns?[{id:'learning',days:learningDays,budget,seasonality:1,learningPct:6},...(settings.days>learningDays?[{id:'steady',days:settings.days-learningDays,budget,seasonality:1,learningPct:0}]:[])]:[{id:'steady',days:settings.days,budget,seasonality:1,learningPct:0}];
 const result=project(base,targets,{name:'Simulação simples',start:settings.start,baselineBudget:settings.reference,elasticity,responsivePct:100,ticket:base.ticket,bureauUnit:base.bureauUnit,conversionPct:0,taxOverride:null,extraFeePct:0,fixedDay:0,cycles});
 // Overrides affect only the hypothetical scenario. The historical control stays intact.
 const hypothetical=settings.orderRate===undefined?base:{...base,ordersDay:base.adsDay*settings.orderRate};
 const proposedRaw=settings.ticket===undefined&&settings.bureauUnit===undefined&&settings.orderRate===undefined?result.scenario:project(hypothetical,targets,{name:'Hipóteses dos cards',start:settings.start,baselineBudget:settings.reference,elasticity:settings.orderRate===undefined?elasticity:1,responsivePct:100,ticket:settings.ticket??base.ticket,bureauUnit:settings.bureauUnit??base.bureauUnit,conversionPct:0,taxOverride:null,extraFeePct:0,fixedDay:0,cycles}).scenario;
 const adjust=(p:typeof result.scenario)=>{const fees=round(p.revenue*settings.feeRate),fixed=round(p.days*settings.fixedDay);return {...p,fees,fixed,profit:round(p.profit-fees-fixed)};};
 const current=adjust(result.control),proposed=adjust(proposedRaw),delta=round(proposed.profit-current.profit),extraAds=round(proposed.ads-current.ads),extraOrders=proposed.orders-current.orders;
 return {budget,elasticity,current,proposed,delta,extraAds,extraOrders,marginalCac:extraAds>0&&extraOrders>0?extraAds/extraOrders:null,unitContribution:(settings.ticket??base.ticket)*(1-(proposed.revenue?proposed.tax/proposed.revenue:0)-settings.feeRate)-(settings.bureauUnit??base.bureauUnit),end:addDays(settings.start,settings.days-1)};
}
export function responsesFor(settings:QuickSettings){return settings.orderRate===undefined?[...quickResponses]:[{...quickResponses[1],label:'Eficiência definida por você'}];}
export function quickCurve(base:Baseline,targets:TaxRow[],settings:QuickSettings){
 const series=responsesFor(settings).map(r=>({...r,points:[...new Set([...Array.from({length:101},(_,i)=>i*100),settings.reference])].sort((a,b)=>a-b).map(budget=>{const p=quickProjection(base,targets,settings,budget,r.elasticity);return {x:budget,y:p.proposed.profit};})}));
 const central=series.find(r=>r.id==='central')!;let best={budget:settings.reference,profit:quickProjection(base,targets,settings,settings.reference,.8).proposed.profit};
 for(const p of central.points)if(p.y>best.profit+.005)best={budget:p.x,profit:p.y};
 const near=best.budget;for(let b=Math.max(0,near-100);b<=Math.min(10000,near+100);b+=10){const p=quickProjection(base,targets,settings,b,.8);if(p.proposed.profit>best.profit+.005)best={budget:b,profit:p.proposed.profit};}
 return {series,best,bounded:best.budget===10000};
}

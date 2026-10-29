import { loadSimulationData, rows } from './simulation-data';
import { dates } from './simulation';
import type { FinanceSnapshot, RefundRow } from './finance';
export async function loadFinanceData(start:string,end:string,signal?:AbortSignal):Promise<FinanceSnapshot>{
 const [snapshot,refunds]=await Promise.all([loadSimulationData(start,end,signal),rows('refunds_daily','date,refund_count,refund_value',`date=gte.${start}&date=lte.${end}&order=date.asc`,signal).then(values=>({values,error:''})).catch(e=>{if(signal?.aborted)throw e;return {values:null,error:'Estornos indisponíveis nesta leitura.'};})]);
 const refundRows:RefundRow[]|null=refunds.values?.map(r=>{const date=String(r.date).slice(0,10);dates(date,date);const count=Number(r.refund_count),value=Number(r.refund_value);if(r.refund_count==null||r.refund_value==null||!Number.isInteger(count)||count<0||!Number.isFinite(value)||value<0)throw new Error('Valor inválido na fonte de estornos.');return {date,refund_count:count,refund_value:value};})??null;
 return {...snapshot,refunds:refundRows,refundWarning:refunds.error};
}

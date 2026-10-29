import { addDays, dates, round, taxAt, type Snapshot } from './simulation';
export const expenseFields = [['fees','Taxas de pagamento'],['people','Equipe e pró-labore'],['systems','Sistemas e serviços'],['other','Outras despesas operacionais'],['financial','Despesas financeiras e outros tributos']] as const;
export type ExpenseKey = typeof expenseFields[number][0];
export type Expenses = Record<ExpenseKey, number | null>;
export type ExpenseMonth = Expenses & {month:string;version:number;updated_at:string;updated_by:string};
export const emptyExpenses = (): Expenses => ({fees:null,people:null,systems:null,other:null,financial:null});
export function expensePayload(values:Record<string,string>):Expenses {
 const out=emptyExpenses();for(const [key,label] of expenseFields){const text=values[key]?.trim()||'';const n=text===''?null:Number(text);if(n!==null&&(!Number.isFinite(n)||n<0||n>1e9||Math.abs(n*100-Math.round(n*100))>1e-5))throw new Error(`Confira ${label}: use um valor positivo ou zero, com até duas casas decimais.`);out[key]=n;}return out;
}
export function monthPeriod(month:string,today:string){
 if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw new Error('Escolha um mês válido.');
 const start=month+'-01';dates(start,start);const next=new Date(start+'T00:00:00Z');next.setUTCMonth(next.getUTCMonth()+1);const last=addDays(next.toISOString().slice(0,10),-1);const end=[last,addDays(today,-1)].sort()[0];if(end<start)throw new Error('Este mês ainda não tem dias encerrados.');return {start,end,calendarDays:Number(last.slice(8)),partial:end<last};
}
export type RefundRow={date:string;refund_count:number;refund_value:number};
export type FinanceSnapshot=Snapshot & {refunds:RefundRow[]|null;refundWarning:string};
export function buildDre(s:FinanceSnapshot,expenses:Expenses=emptyExpenses()){
 const calendar=dates(s.start,s.end),inside=(r:{date:string})=>r.date>=s.start&&r.date<=s.end;
 const revenue=s.revenue.filter(inside),bureau=s.bureau.filter(inside),ads=s.ads.filter(inside),issues:string[]=[];
 for(const [label,rows,key] of [['receita',revenue,(r:{date:string})=>r.date],['bureau',bureau,(r:{date:string})=>r.date],['Ads',ads,(r:{date:string;campaign_id?:string})=>r.date+':'+r.campaign_id]] as const){if(new Set(rows.map(key)).size!==rows.length)issues.push(`Duplicação na fonte de ${label}.`);if(calendar.some(day=>!rows.some(r=>r.date===day)))issues.push(`A fonte de ${label} não cobre todos os dias; ausência não foi tratada como zero.`);}
 if(bureau.some(r=>!Number.isFinite(r.custo_bureau)||r.custo_bureau<0)||ads.some(r=>!Number.isFinite(r.cost_micros)||r.cost_micros<0)||revenue.some(r=>!Number.isFinite(r.revenue)||r.revenue<0||!Number.isInteger(r.paid_orders)||r.paid_orders<0))issues.push('Valores inválidos nas fontes financeiras.');
 const sum=<T,>(rows:T[],fn:(r:T)=>number)=>round(rows.reduce((n,r)=>n+fn(r),0));
 const sales=sum(revenue,r=>r.revenue),orders=sum(revenue,r=>r.paid_orders),tax=sum(revenue,r=>r.revenue*taxAt(s.targets,r.date).value/100),consultations=sum(bureau,r=>r.custo_bureau),media=sum(ads,r=>r.cost_micros/1e6),netRevenue=round(sales-tax),grossProfit=round(netRevenue-consultations),panelProfit=round(grossProfit-media);
 const missing=expenseFields.filter(([key])=>expenses[key]===null).map(([,label])=>label);
 for(const [key] of expenseFields)if(expenses[key]!==null&&(!Number.isFinite(expenses[key])||expenses[key]!<0||expenses[key]!>1e9))issues.push('Despesa inválida.');
 const complements=round(expenseFields.reduce((n,[key])=>n+(expenses[key]??0),0)),profit=round(panelProfit-complements);
 const refunds=s.refunds===null?null:sum(s.refunds.filter(inside),r=>r.refund_value);
 return {days:calendar.length,sales,orders,tax,consultations,media,netRevenue,grossProfit,panelProfit,complements,profit,margin:sales>0?profit/sales*100:null,missing,complete:!missing.length&&!s.warnings.length&&!issues.length,issues,refunds,expenses:{...expenses}};
}

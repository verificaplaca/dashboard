export const overviewPresets=[['today','Hoje'],['yesterday','Ontem'],['week','Esta semana'],['7d','Últimos 7 dias'],['14d','Últimos 14 dias'],['30d','Últimos 30 dias'],['60d','Últimos 60 dias'],['90d','Últimos 90 dias'],['month','Este mês'],['lastmonth','Mês passado'],['ytd','Este ano até agora'],['custom','Personalizado']] as const;
export type OverviewPreset=typeof overviewPresets[number][0];
export interface OverviewPeriod {preset:OverviewPreset;from?:string;to?:string}
export function validateOverviewPeriod(period:OverviewPeriod){
 if(!overviewPresets.some(([key])=>key===period.preset))throw new Error('Escolha um período válido.');
 if(period.preset==='custom'){
  const valid=(day?:string)=>!!day&&/^\d{4}-\d{2}-\d{2}$/.test(day)&&Number.isFinite(Date.parse(day))&&new Date(day).toISOString().slice(0,10)===day;
  if(!valid(period.from)||!valid(period.to))throw new Error('Informe as datas inicial e final.');
  if(period.from!>period.to!)throw new Error('A data inicial deve ser anterior ou igual à final.');
 }
 return period;
}
export function readOverviewPeriod(storage:Pick<Storage,'getItem'>):OverviewPeriod{
 try{const saved=JSON.parse(storage.getItem('vp_filter')||'null');if(saved)return validateOverviewPeriod(saved);}catch{}
 return {preset:'month'};
}

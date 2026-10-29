import type { SupabaseClient } from '@supabase/supabase-js';
import type { Expenses, ExpenseMonth } from './finance';
export class FinanceConflict extends Error {}
export async function readExpenseMonth(client:SupabaseClient,month:string){
 const access=await client.rpc('console_finance_access');if(access.error)throw new Error('Não foi possível verificar o acesso às despesas.');if(access.data!==true)return {allowed:false,row:null};
 const result=await client.from('console_finance_months').select('*').eq('month',month+'-01').maybeSingle();if(result.error)throw new Error('Não foi possível carregar as despesas do mês.');return {allowed:true,row:result.data as ExpenseMonth|null};
}
export async function saveExpenseMonth(client:SupabaseClient,month:string,values:Expenses,version:number):Promise<ExpenseMonth>{
 const result=await client.rpc('console_save_finance_month',{p_month:month+'-01',p_values:values,p_version:version});
 if(result.error?.code==='40001')throw new FinanceConflict('Outra pessoa atualizou este mês. Seu rascunho está preservado. Carregue os valores atuais antes de salvar.');
 if(result.error)throw new Error(result.error.code==='42501'?'Sua conta não tem permissão para salvar despesas.':'Não foi possível salvar. Confira os valores e tente novamente.');
 return result.data as ExpenseMonth;
}

import { overviewUnavailableMessage } from './domain/overview-publication';
import { useEffect, useRef, useState } from 'react';
import { Activity, ArrowRight, ArrowUpRight, CalendarDays, ChartNoAxesCombined, CheckCircle2, CircleDollarSign, Gauge, LayoutDashboard, Megaphone, RefreshCw, ShoppingBag, Target, Users, Wallet } from 'lucide-react';
import { useAuth } from './Auth';
import { go } from './format';
import { greeting, saoPauloClock } from './domain/clock';
import { addDays } from './domain/simulation';
import { dispatchSummary, type HomeDispatch } from './domain/home-overview';
import OriginalOverview, {type OverviewEvent, type OverviewHandle} from './OriginalOverview';
import {homeOverviewValues, type OverviewHomeSnapshot} from './domain/overview-home-snapshot';
import { useGlobalLoading } from './Loading';
import { readPublic } from './api';
import { Dialog } from './components';
import {overviewPresets,validateOverviewPeriod,type OverviewPeriod} from './domain/overview-period';
import TaskHomeSummary from './TaskHomeSummary';
import './operations-home.css';
const money=(value:number|null|undefined)=>value==null?'—':value.toLocaleString('pt-BR',{style:'currency',currency:'BRL',maximumFractionDigits:2});
const number=(value:number|null|undefined)=>value==null?'—':value.toLocaleString('pt-BR');
const shortDate=(day:string)=>new Date(day+'T12:00:00Z').toLocaleDateString('pt-BR',{day:'2-digit',month:'2-digit',timeZone:'America/Sao_Paulo'});
function useHomeSignals(from:string,revision:number){
 const [state,setState]=useState<{loading:boolean;dispatches:HomeDispatch[]|null;trackingError:string}>({loading:true,dispatches:null,trackingError:''});
 useGlobalLoading(state.loading);
 useEffect(()=>{
  const abort=new AbortController();setState(previous=>({...previous,loading:true}));
  void Promise.allSettled([
   readPublic<HomeDispatch>('conversion_dispatches_public',new URLSearchParams({select:'checkout_id,ga4_status,ads_status,meta_status,created_at',created_at:`gte.${from}T03:00:00Z`,order:'created_at.desc,checkout_id.asc',limit:'50'}),abort.signal),
  ]).then(([dispatches])=>{if(!abort.signal.aborted)setState({loading:false,dispatches:dispatches.status==='fulfilled'?dispatches.value:null,trackingError:dispatches.status==='rejected'?'Não foi possível consultar os envios.':''});});
  return()=>abort.abort();
 },[from,revision]);
 return state;
}
export default function OperationsHome(){
 const {user,role}=useAuth(),[now,setNow]=useState(()=>new Date()),[revision,setRevision]=useState(0);
 useEffect(()=>{const timer=setInterval(()=>setNow(new Date()),60_000);return()=>clearInterval(timer);},[]);
 const clock=saoPauloClock(now),engine=useRef<OverviewHandle>(null);
 const [source,setSource]=useState<{snapshot:OverviewHomeSnapshot|null;period:string;loading:boolean;error:string}>({snapshot:null,period:'',loading:true,error:''});
 const signals=useHomeSignals(addDays(clock.day,-6),revision);
 const summary=homeOverviewValues(source.snapshot),valid=!!summary;
 const end=source.snapshot?.days.at(-1)?.date||source.snapshot?.filter?.to||clock.day,month=end.slice(0,7),monthCount=new Set(source.snapshot?.days.map(day=>day.date.slice(0,7))).size,singleMonth=monthCount===1;
 const targets=source.snapshot?.real&&monthCount>0?source.snapshot.targets:null,revenueGoal=targets?.revenue??null,cacGoal=targets?.cpa??null,budget=targets?.budget??null;
 const cac=summary?.cac??null,progress=summary&&singleMonth&&revenueGoal&&revenueGoal>0?summary.sales/revenueGoal*100:null;
 const dispatches=signals.dispatches?dispatchSummary(signals.dispatches):null;
 const days=valid?source.snapshot!.days.filter(row=>row.date>=addDays(end,-6)&&row.date<=end):[],max=Math.max(1,...days.map(row=>row.revenue))*1.2;
 const includesToday=!!source.snapshot?.days.some(day=>day.date===clock.day);
 const name=typeof user?.user_metadata?.full_name==='string'?user.user_metadata.full_name.trim().split(' ')[0]:'';
 const busy=source.loading||signals.loading;
 const receiveOverview=(event:OverviewEvent)=>setSource(previous=>{
  if(event.snapshot)return {...previous,snapshot:event.snapshot.real?event.snapshot:null,period:event.period??previous.period,loading:event.loading??false,error:event.snapshot.real?'':overviewUnavailableMessage};
  if(event.error)return {...previous,snapshot:null,error:overviewUnavailableMessage,loading:false};
  if(typeof event.loading==='boolean'&&event.loading!==previous.loading)return {...previous,loading:event.loading};
  return previous;
 });
 const [periodEditor,setPeriodEditor]=useState(false);
 const refresh=()=>{engine.current?.refresh();setRevision(n=>n+1);};
 const modules=[{path:'/dashboard',title:'Visão Geral',copy:'Indicadores e análises',icon:LayoutDashboard},{path:'/financeiro',title:'Financeiro',copy:'Resultado e despesas',icon:Wallet},{path:'/simulador',title:'Simulador',copy:'Cenários de investimento',icon:Gauge},{path:'/marketing',title:'Campanhas',copy:'Aquisição e mídia',icon:Megaphone},{path:'/metas',title:'Metas',copy:'Objetivos do negócio',icon:Target},{path:'/tracking',title:'Tracking',copy:'Qualidade dos envios',icon:Activity},{path:'/tarefas',title:'Tarefas',copy:'Execução da equipe',icon:CheckCircle2},...(role==='administrator'?[{path:'/usuarios',title:'Usuários',copy:'Perfis e acessos',icon:Users}]:[])];
 return <div className="overview-home">
  <OriginalOverview ref={engine} summary dataOnly onEvent={receiveOverview}/>
  <header className="oh-heading"><div><span className="oh-eyebrow">CENTRAL DA OPERAÇÃO</span><h1>{greeting(clock.hour)}{name?`, ${name}`:''}.</h1><p>O essencial para acompanhar o negócio e decidir o próximo passo.</p></div><div className="oh-heading-actions"><span className="oh-date"><CalendarDays size={15}/>{now.toLocaleDateString('pt-BR',{day:'numeric',month:'long',timeZone:'America/Sao_Paulo'})}</span><button className="button secondary" disabled={busy||!source.snapshot} onClick={()=>setPeriodEditor(true)}><CalendarDays size={15}/>{overviewPresets.find(([key])=>key===source.snapshot?.filter?.preset)?.[1]||'Este mês'}</button><button className="button secondary" disabled={busy} onClick={refresh}><RefreshCw size={15}/>Atualizar resumo</button></div></header>
  <div className="oh-section-line"><h2>Seu negócio, em um olhar</h2><span>Visão Geral · {source.period||'Período atual'}{includesToday?' · inclui hoje':''}</span></div>
  {source.error?<div className="oh-notice" role="status">{source.error} <button onClick={refresh}>Tentar novamente</button></div>:null}
  <section className="oh-kpis" aria-label="Indicadores essenciais do período">
   {[
    {label:'Receita Bruta',value:money(summary?.sales),detail:'Pedidos pagos válidos',icon:CircleDollarSign,path:'/dashboard',kind:''},
    {label:'Lucro Líquido',value:money(summary?.profit),detail:'Após imposto, Ads e bureau',icon:Wallet,path:'/dashboard',kind:summary?(summary.profit<0?'negative':'positive'):''},
    {label:'Transações',value:number(summary?.orders),detail:'Vendas confirmadas no período',icon:ShoppingBag,path:'/dashboard',kind:''},
    {label:'CAC',value:money(cac),detail:cacGoal===null?'Google Ads ÷ pedidos pagos':`Meta CAC: ${money(cacGoal)}`,icon:Target,path:'/marketing',kind:cac!==null&&cacGoal!==null&&cac>cacGoal?'negative':''},
   ].map(item=><button className={`oh-kpi ${item.kind}`} key={item.label} onClick={()=>go(item.path)}><span className="oh-kpi-label">{item.label}<item.icon size={18}/></span><strong>{item.value}</strong><span className="oh-kpi-detail">{item.detail}<ArrowUpRight size={14}/></span></button>)}
  </section>
  <div className="oh-financial-grid">
   <section className="oh-panel oh-revenue"><div className="oh-panel-heading"><div><span className="oh-eyebrow">RITMO DAS VENDAS</span><h2>Como a receita evoluiu</h2></div><button className="oh-link" onClick={()=>go('/dashboard')}>Ver análise<ArrowUpRight size={15}/></button></div><div className="oh-chart-caption"><span><i/>Receita diária</span><small>{days.length?`${days.length} ${days.length===1?'dia recente':'dias recentes'} do período`:'Sem dias no período'}</small></div>
    {days.length?<div className="oh-bars" style={{gridTemplateColumns:`repeat(${days.length},minmax(0,1fr))`}} role="img" aria-label={days.map(day=>`${shortDate(day.date)}: ${money(day.revenue)}`).join('; ')}>{days.map(day=><div className="oh-bar-column" key={day.date} title={`${shortDate(day.date)}: ${money(day.revenue)}`}><span className="oh-bar-value">{day.revenue.toLocaleString('pt-BR',{notation:'compact',maximumFractionDigits:1})}</span><div className="oh-bar-track"><div style={{height:`${Math.max(day.revenue>0?3:0,day.revenue/max*100)}%`}}/></div><span className="oh-bar-date">{shortDate(day.date)}</span></div>)}</div>:<div className="oh-unavailable">{valid?'Sem dados no período selecionado.':'A evolução aparece após uma leitura válida das fontes.'}</div>}
    <div className="oh-chart-footer"><ChartNoAxesCombined size={16}/><span>{includesToday?'Inclui hoje: os valores do dia ainda estão em atualização.':'Mesmo recorte de datas da Visão Geral.'}</span></div>
   </section>
   <section className="oh-panel oh-goal"><div className="oh-panel-heading"><div><span className="oh-eyebrow">OBJETIVO DO MÊS</span><h2>Meta de receita</h2></div><Target size={20}/></div>
    <div className="oh-goal-center"><div className="oh-ring"><svg viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="42"/><circle cx="50" cy="50" r="42" strokeDasharray={2*Math.PI*42} strokeDashoffset={2*Math.PI*42*(1-Math.min(100,Math.max(0,progress??0))/100)}/></svg><strong>{progress===null?'—':`${progress.toLocaleString('pt-BR',{maximumFractionDigits:1})}%`}</strong></div><span>{progress===null?(revenueGoal===null?'Meta não disponível':revenueGoal===0?'Meta zerada':monthCount===0?'Sem dias no período':!singleMonth?'Período com mais de um mês':'Receita indisponível'):progress>=100?'Meta atingida':'Do objetivo mensal realizado'}</span></div>
    <div className="oh-goal-detail"><span>Meta de {new Date(month+'-01T12:00:00Z').toLocaleDateString('pt-BR',{month:'long',timeZone:'UTC'})}</span><strong>{money(revenueGoal)}</strong></div><p>{(revenueGoal===0?'Defina uma meta maior que zero para acompanhar o progresso.':progress===null?(!singleMonth?'Selecione um único mês na Visão Geral para acompanhar o progresso mensal.':'Atualize a Visão Geral para acompanhar o progresso.'):`Receita realizada até ${shortDate(end)}. A meta considera o mês completo.`)}</p><button className="oh-link" onClick={()=>go('/metas')}>{role==='administrator'?'Revisar metas':'Consultar metas'}<ArrowRight size={15}/></button>
   </section>
  </div>
  <div className="oh-execution-grid"><div className="oh-task-panel"><TaskHomeSummary revision={revision}/></div>
   <section className="oh-panel oh-channels"><div className="oh-panel-heading"><div><span className="oh-eyebrow">CAMPANHAS E ENTREGAS</span><h2>Os canais da operação</h2></div><Activity size={20}/></div>
    <button className="oh-channel" onClick={()=>go('/marketing')}><span className="oh-channel-icon"><Megaphone size={20}/></span><span className="oh-channel-body"><strong>Google Ads</strong><small>Investimento no período da Visão Geral</small></span><span className="oh-channel-value">{money(summary?.media)}<ArrowUpRight size={15}/></span></button>
    <div className="oh-channel-note">{(budget===null?'Orçamento mensal não cadastrado.':`Orçamento mensal: ${money(budget)}.`)}</div>
    <button className="oh-channel" onClick={()=>go('/tracking')}><span className="oh-channel-icon"><Activity size={20}/></span><span className="oh-channel-body"><strong>Tracking de conversões</strong><small>{dispatches?`${dispatches.total} envios consultados nos últimos 7 dias`:'Consulta dos envios indisponível'}</small></span><ArrowUpRight size={16}/></button>
    {dispatches?<div className="oh-delivery-stats"><div className={dispatches.failed?'warning':''}><strong>{dispatches.failed}</strong><span>Com falha</span></div><div><strong>{dispatches.pending}</strong><span>Pendentes</span></div><div><strong>{dispatches.unknown}</strong><span>A conferir</span></div></div>:<p className="oh-source-note">{signals.trackingError}</p>}
    <p className="oh-source-note">Amostra dos últimos 50 envios em 7 dias. Consulte cada canal no Tracking.</p>
    <button className="oh-finance-note" onClick={()=>go('/financeiro')}><span><Wallet size={17}/></span><span>Fechamento e despesas na DRE<small>O lucro da Visão Geral não inclui as despesas complementares do Financeiro.</small></span><ArrowRight size={15}/></button>
   </section>
  </div>
  <section className="oh-modules"><div className="oh-section-line"><h2>Explore a operação</h2><span>Um caminho direto para cada área</span></div><nav className="oh-module-grid" aria-label="Acessos rápidos da operação">{modules.map(module=><button key={module.path} onClick={()=>go(module.path)}><module.icon size={20}/><span><strong>{module.title}</strong><small>{module.copy}</small></span><ArrowUpRight size={15}/></button>)}</nav></section>
  <p className="oh-footnote">Indicadores, metas e período vêm do mesmo motor da Visão Geral. O filtro de período é compartilhado com a Visão Geral. Tarefas e envios têm os escopos indicados em cada bloco.</p>
 {periodEditor&&<HomePeriodDialog initial={source.snapshot?.filter??{preset:'month'}} today={clock.day} close={()=>setPeriodEditor(false)} apply={period=>{engine.current?.setPeriod(period);setPeriodEditor(false);}}/>}
 </div>;
}
function HomePeriodDialog({initial,today,close,apply}:{initial:OverviewPeriod;today:string;close:()=>void;apply:(period:OverviewPeriod)=>void}){
 const [preset,setPreset]=useState(initial.preset),[error,setError]=useState('');
 return <Dialog title="Período do resumo" close={close}><form className="live-form" onSubmit={event=>{event.preventDefault();const values=new FormData(event.currentTarget);try{apply(validateOverviewPeriod({preset,from:String(values.get('from')||''),to:String(values.get('to')||'')}));}catch(reason){setError((reason as Error).message);}}}><div className="form-fields"><p>As datas escolhidas também serão usadas na Visão Geral.</p><label>Período<select value={preset} onChange={event=>setPreset(event.target.value as OverviewPeriod['preset'])}>{overviewPresets.map(([key,label])=><option key={key} value={key}>{label}</option>)}</select></label>{preset==='custom'&&<><label>Data inicial<input name="from" type="date" required defaultValue={initial.from||today.slice(0,7)+'-01'}/></label><label>Data final<input name="to" type="date" required defaultValue={initial.to||today}/></label></>}{error&&<p role="alert">{error}</p>}</div><div className="dialog-footer"><button type="button" className="button secondary" onClick={close}>Cancelar</button><button className="button primary">Aplicar período</button></div></form></Dialog>;
}

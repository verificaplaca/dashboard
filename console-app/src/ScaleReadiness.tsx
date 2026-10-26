import { Check, Minus, ShieldCheck } from 'lucide-react';
import { addDays } from './domain/simulation';
import { SCALE_LEVELS, type ScaleDecision, type ScaleState } from './domain/scale-plan';
const money = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const day = (s: string) => s.split('-').reverse().join('/');
export default function ScaleReadiness({ decision, state, loading, error, today }: { decision: ScaleDecision; state: ScaleState; loading: boolean; error: string; today: string }) {
  const entry = state.entries.at(-1)!, initial = entry.level === 0;
  const firstReview = addDays(entry.changedOn, decision.required + state.lagDays + 1);
  const ready = !loading && !error && decision.valid;
  const heading = !ready ? 'Aguardando validação dos dados reais' : decision.action === 'reduce' ? 'Conter a escala e revisar a etapa' : initial ? decision.action === 'increase' ? 'Critérios atendidos para testar R$ 3.100/dia' : 'Manter R$ 2.790/dia antes do próximo teste' : `Avaliar a etapa registrada de ${money(SCALE_LEVELS[entry.level])}/dia`;
  return <section className="sim-readiness" aria-labelledby="sim-readiness-title">
    <div className="sim-growth-heading"><div><span className="sim-data-kind real">RESULTADO REAL · PRÓXIMO PASSO</span><h3 id="sim-readiness-title">{heading}</h3><p>Os critérios abaixo usam CAC e lucro realizados. A escolha de base, elasticidade e aprendizado não altera esta avaliação.</p></div><ShieldCheck size={22} aria-hidden="true" /></div>
    {!ready ? <p className="sim-alert" role="status">{loading ? 'Atualizando as fontes originais…' : error || decision.reasons.join(' ') || 'Atualize a leitura antes de decidir.'} Nenhum aumento liberado nesta leitura.</p> : <>
      <p className="sim-readiness-summary">{decision.action === 'increase' && decision.nextLevel !== null ? `Próximo teste: ${money(SCALE_LEVELS[decision.nextLevel])}/dia. Registre apenas depois de aplicar a mudança na conta.` : decision.action === 'hold' && initial ? 'Próximo teste: R$ 3.100/dia. Ainda não liberado; confira os requisitos pendentes abaixo.' : decision.reasons.join(' ')}</p>
      <div className="sim-scale-checks sim-readiness-checks">{decision.checks.map(c => <div key={c.label}><span className={c.pass ? 'pass' : ''}>{c.pass ? <Check size={13} aria-hidden="true" /> : <Minus size={13} aria-hidden="true" />}</span><div><strong>{c.pass ? 'Atendido' : 'Pendente'} · {c.label}</strong><p>{c.detail}</p></div></div>)}</div>
      <p className="sim-help">{today < firstReview ? `Primeira reavaliação possível: ${day(firstReview)}, com ${decision.required} dias completos e a reserva de atualização registrada.` : `A data mínima de ${day(firstReview)} já passou; os demais critérios continuam obrigatórios.`} Completar a janela não libera o aumento sozinho.</p>
    </>}
    <p className="sim-help"><strong>Para o primeiro teste de R$ 3.100:</strong> janela válida e completa, CAC ponderado ≤ R$ 9, lucro/dia positivo, acumulado da etapa positivo e gasto compatível com a verba registrada (±20%). Sem sinal de redução. Se qualquer requisito faltar, manter e reavaliar. Nos aumentos seguintes, exigir também lucro/dia pelo menos 5% maior, gasto efetivo maior e transição recuperada contra a referência validada.</p>
    {!initial && <p className="sim-help">R$ 2.790 é o controle das projeções. A etapa local já está em {money(SCALE_LEVELS[entry.level])}/dia; este quadro respeita o histórico registrado, sem sugerir voltar automaticamente à verba inicial.</p>}
    <p className="sim-help">O mínimo de 14 dias e os limites de CAC/lucro são regras deste plano; o ciclo de conversão e a reserva são ajustáveis em Decisão de hoje. Este módulo não altera campanhas.</p>
  </section>;
}

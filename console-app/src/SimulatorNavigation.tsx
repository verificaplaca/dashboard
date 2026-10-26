import { Activity, ChartNoAxesCombined, SlidersHorizontal, TrendingUp } from 'lucide-react';
export type SimulatorView = 'decision' | 'plan' | 'optimizer' | 'custom';
export const SIMULATOR_VIEWS = [
  { id: 'decision', label: 'Decisão de hoje', hint: 'CAC e lucro realizados', icon: Activity },
  { id: 'plan', label: 'Plano até janeiro', hint: 'Bases e critérios de escala', icon: ChartNoAxesCombined },
  { id: 'optimizer', label: 'Investimento ideal', hint: 'Testar bases de 7, 14 e 30 dias', icon: TrendingUp },
  { id: 'custom', label: 'Cenário livre', hint: 'Editar, salvar e reabrir hipóteses', icon: SlidersHorizontal },
] as const;
export default function SimulatorNavigation({ view, onChange }: { view: SimulatorView; onChange: (view: SimulatorView) => void }) {
  return <div className="sim-navigation" role="tablist" aria-label="Áreas do simulador" onKeyDown={e => {
    const index = SIMULATOR_VIEWS.findIndex(v => v.id === view);
    const next = e.key === 'ArrowRight' ? (index + 1) % 4 : e.key === 'ArrowLeft' ? (index + 3) % 4 : e.key === 'Home' ? 0 : e.key === 'End' ? 3 : -1;
    if (next < 0) return; e.preventDefault(); onChange(SIMULATOR_VIEWS[next].id); (e.currentTarget.querySelectorAll('button')[next] as HTMLButtonElement).focus();
  }}>{SIMULATOR_VIEWS.map(({ id, label, hint, icon: Icon }) => <button key={id} type="button" id={`sim-tab-${id}`} role="tab" aria-selected={view === id} aria-controls={`sim-panel-${id}`} tabIndex={view === id ? 0 : -1} onClick={() => onChange(id)}><Icon size={17} /><span><strong>{label}</strong><small>{hint}</small></span></button>)}</div>;
}

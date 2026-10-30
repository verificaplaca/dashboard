import { useId, useState, type CSSProperties } from 'react';

export interface FunnelStage { label: string; value: number; color: string; endColor: string }
const number = (value: number) => value.toLocaleString('pt-BR');
const percent = (value: number) => `${value.toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%`;

// Curved segment geometry from the supplied design, with three translucent rings.
function segment(start: number, end: number, layer: number) {
  const a = start * 88 * layer, b = end * 88 * layer;
  return `M0 ${100-a} C55 ${100-a},45 ${100-b},100 ${100-b} L100 ${100+b} C45 ${100+b},55 ${100+a},0 ${100+a} Z`;
}

/** CSS is scoped to this component so it also works inside the overview ShadowRoot. */
export function FunnelChart({ data }: { data: FunnelStage[] }) {
  const id = useId().replace(/:/g, '');
  const [hovered, setHovered] = useState<number | null>(null);
  const base = data[0]?.value ?? 0;
  const max = Math.max(1, ...data.map(stage => stage.value));
  const hasData = data.some(stage => stage.value > 0);
  return <div className="flow-funnel">
    <div className="flow-funnel-stages" aria-label="Etapas do funil de conversão">
      {data.map((stage, index) => {
        const next = data[Math.min(index + 1, data.length - 1)];
        const pct = base > 0 ? percent(stage.value / base * 100) : '—';
        return <div className="flow-funnel-stage" key={stage.label} tabIndex={0}
          aria-label={`${stage.label}: ${number(stage.value)}. ${base > 0 ? `${pct} dos checkouts` : 'Percentual indisponível: sem checkouts'}.`}
          data-dimmed={hovered !== null && hovered !== index} data-active={hovered === index}
          onMouseEnter={() => setHovered(index)} onMouseLeave={() => setHovered(null)}
          onFocus={() => setHovered(index)} onBlur={() => setHovered(null)}
          style={{ '--stage-delay': `${index * 120}ms` } as CSSProperties}>
          <strong className="flow-funnel-value">{number(stage.value)}</strong>
          <svg viewBox="0 0 100 200" preserveAspectRatio="none" aria-hidden="true">
            <defs><linearGradient id={`${id}-${index}`} x1="0" x2="1" y1="0" y2="0">
              <stop offset="0%" stopColor={stage.color}/><stop offset="100%" stopColor={stage.endColor}/>
            </linearGradient></defs>
            {[1, 1 - 1 / 3 * .35, 1 - 2 / 3 * .35].map((scale, ring) =>
              <path key={ring} d={segment(stage.value / max, next.value / max, scale)}
                fill={ring === 2 ? `url(#${id}-${index})` : stage.color} opacity={.18 + ring * .325}
                style={{ '--ring-scale': 1 + ring * .06 } as CSSProperties}/>
            )}
          </svg>
          <span className="flow-funnel-percent">{pct}</span>
          <span className="flow-funnel-label">{stage.label}</span>
        </div>;
      })}
    </div>
    <p className="flow-funnel-caption">{hasData ? 'Percentuais no gráfico: participação sobre os checkouts.' : 'Sem dados de conversão no período selecionado.'}</p>
    <div className="flow-funnel-rates">
      {data.slice(1).map((stage, index) => <div key={stage.label}>
        <span>{data[index].label} → {stage.label}</span>
        <strong>{data[index].value > 0 ? percent(stage.value / data[index].value * 100) : '—'}</strong>
      </div>)}
    </div>
  </div>;
}

import { useState } from 'react';
import { ArrowUpRight, ShieldCheck } from 'lucide-react';
import { go } from './format';
import OriginalOverview from './OriginalOverview';

/** Reuses the original financial engine; no second financial calculation model. */
export default function ExecutiveSummary() {
  const [example, setExample] = useState(false);
  return <section className="executive-summary" aria-labelledby="executive-summary-title">
    <div className="home-section-heading"><div><span className="home-kicker"><ShieldCheck size={13} />{example ? 'EXEMPLO · FONTE INDISPONÍVEL' : 'FONTES ORIGINAIS'}</span><h2 id="executive-summary-title">Resumo executivo</h2><p>Receita, custos e resultado, com as mesmas regras da Visão Geral.</p></div><button className="button secondary" onClick={() => go('/dashboard')}>Abrir Visão Geral<ArrowUpRight size={15} /></button></div>
    <OriginalOverview summary onEvent={event => { if (typeof event.example === 'boolean') setExample(event.example); }} />
  </section>;
}

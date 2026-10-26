import { useState, type ReactNode } from 'react';
/** Render detailed charts on first expansion; retain their state on subsequent visits. */
export default function SimulatorDisclosure({ title, children }: { title: string; children: ReactNode }) {
  const [visited, setVisited] = useState(false);
  return <details className="sim-ux-disclosure" onToggle={e => { if (e.currentTarget.open) setVisited(true); }}><summary>{title}</summary>{visited && children}</details>;
}

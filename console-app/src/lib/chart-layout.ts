import { useEffect, useState } from 'react';

/** SVG charts need a smaller coordinate system to keep text readable on phones. */
export function useCompactChart() {
  const [compact, setCompact] = useState(() => matchMedia('(max-width:600px)').matches);
  useEffect(() => {
    const query = matchMedia('(max-width:600px)');
    const update = () => setCompact(query.matches);
    update(); query.addEventListener('change', update);
    return () => query.removeEventListener('change', update);
  }, []);
  return compact;
}

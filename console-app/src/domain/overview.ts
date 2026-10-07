/** Dedicated reader; never replaces the application's global fetch. */
export function readOnlyOverviewFetch(input: RequestInfo | URL, options: RequestInit = {}, signal?: AbortSignal) {
  const url = new URL(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url, location.href);
  const method = (options.method || (input instanceof Request ? input.method : 'GET')).toUpperCase();
  if (method !== 'GET' || url.protocol !== 'https:' || (url.port && url.port !== '443') || url.username || url.password || !['ftmgmfdqdqxboiktxcoj.supabase.co', 'ozquoloetuzynnyzkado.supabase.co'].includes(url.hostname) || !url.pathname.startsWith('/rest/v1/')) return Promise.reject(new Error('A visão geral integrada permite somente leitura das fontes originais.'));
  return fetch(input, { ...options, ...(signal ? { signal } : {}) });
}

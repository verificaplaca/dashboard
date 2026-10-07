(() => {
  const request = window.fetch.bind(window);
  const allowed = new Set(['ftmgmfdqdqxboiktxcoj.supabase.co', 'ozquoloetuzynnyzkado.supabase.co']);
  // This embedded overview reads existing public aggregates only. Authentication
  // and monthly-target editing belong to the original application's workflow.
  window.fetch = (input, options = {}) => {
    const url = new URL(typeof input === 'string' ? input : input.url, location.href);
    const method = String(options.method || (input instanceof Request ? input.method : 'GET')).toUpperCase();
    if (method !== 'GET' || !allowed.has(url.hostname) || !url.pathname.startsWith('/rest/v1/')) {
      return Promise.reject(new Error('A visão geral integrada permite somente leitura das fontes originais.'));
    }
    return request(input, options);
  };
})();

import fs from 'node:fs';
import crypto from 'node:crypto';
// Frozen upstream source. Only presentation and embedded navigation are adapted.
// The entire calculation/loading script is retained byte-for-byte.
const original = fs.readFileSync(new URL('../reference/dashboard-original.html', import.meta.url), 'utf8');
let html = original.replace('href="favicon-verifica-placa-2.jpg"', 'href="https://verificaplaca.github.io/dashboard/favicon-verifica-placa-2.jpg"');
html = html.replace("try { savedView = localStorage.getItem('vp_sidebar_view') || 'overview'; } catch (e) {}", "// Embedded overview always opens in its own view; shell owns navigation.");
html = html.replace('<head>', '<head>\n<script src="./read-only.js"></script>');
html = html.replace('  <style>', '<script src="./chart-presentation.js"></script>\n  <style>');
html = html.replace('</head>', '<link rel="stylesheet" href="./overview.css" />\n</head>');
for (const [label, id] of [['Resultados Financeiros', 'financial'], ['Por Produto — Teste Pacote Completo', 'products'], ['Monitoramento Operacional', 'operation'], ['Performance Diária', 'performance']]) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  html = html.replace(new RegExp(`(<div class="section-label"[^>]*)(>${escaped})`), `$1 id="section-${id}"$2`);
}
html = html.replace('</body>', '<script src="./bridge.js"></script>\n</body>');
fs.writeFileSync(new URL('../public/overview/index.html', import.meta.url), html);
fs.writeFileSync(new URL('../reference/provenance.json', import.meta.url), JSON.stringify({ url: 'https://verificaplaca.github.io/dashboard/dashboard.html', captured: '2026-10-07', sha256: crypto.createHash('sha256').update(original).digest('hex'), policy: 'Same calculation/loading script; local presentation; read-only requests.' }, null, 2) + '\n');
console.log('Overview prepared from preserved original source.');

// Native overview: same engine and DOM, with no embedded browsing context.
const calculationScript = [...original.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)].filter(m => !/\bsrc\s*=/.test(m[1])).map(m => m[2]).find(s => s.includes('function totals()') && s.includes('async function loadSupabaseData()'));
const body = html.match(/<body[^>]*>([\s\S]*)<\/body>/i)[1].replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '').replace(/<iframe\b[^>]*>[\s\S]*?<\/iframe>/gi, '');
const styles = [...html.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi)].map(m => m[1]).join('\n') + '\n' + fs.readFileSync(new URL('../public/overview/overview.css', import.meta.url), 'utf8');
fs.writeFileSync(new URL('../public/overview/model.json', import.meta.url), JSON.stringify({ body, styles, calculationScript }));

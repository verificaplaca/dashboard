import fs from 'node:fs';
import { execFileSync } from 'node:child_process';

// Build only the browser app. Database migrations and ingestion jobs stay separate.
execFileSync('npm', ['run', 'build'], {
  cwd: new URL('../', import.meta.url),
  stdio: 'inherit',
  env: { ...process.env, VITE_ORIGINAL_DASHBOARD_URL: './dashboard.html' },
});

// Keep the original dashboard and its two embedded panels on the same host.
// Use an explicit list so server scripts, credentials and source are never served.
for (const file of [
  'dashboard.html',
  'google-ads-cac.html',
  'google-ads-cac-service.js',
  'tracking-gateway.html',
  'favicon-verifica-placa-2.jpg',
  'logotipo-verifica-placa.svg',
]) {
  fs.copyFileSync(new URL(`../../${file}`, import.meta.url), new URL(`../dist/${file}`, import.meta.url));
}
fs.writeFileSync(new URL('../dist/robots.txt', import.meta.url), 'User-agent: *\nDisallow: /\n');
console.log('Vercel output ready: console at / and original dashboard at /dashboard.html.');

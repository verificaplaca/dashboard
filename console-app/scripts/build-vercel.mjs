import fs from 'node:fs';
import { execFileSync } from 'node:child_process';

// Build only the browser app. Database migrations and ingestion jobs stay separate.
execFileSync('npm', ['run', 'build'], {
  cwd: new URL('../', import.meta.url),
  stdio: 'inherit',
});

// Publish shared images only; retired dashboards are redirected by Vercel.
// Use an explicit list so server scripts, credentials and source are never served.
for (const file of [
  'favicon-verifica-placa-2.jpg',
  'logotipo-verifica-placa.svg',
]) {
  fs.copyFileSync(new URL(`../../${file}`, import.meta.url), new URL(`../dist/${file}`, import.meta.url));
}
// The standalone overview is a build reference, not a second public application.
fs.rmSync(new URL('../dist/overview/index.html', import.meta.url), { force: true });
fs.writeFileSync(new URL('../dist/robots.txt', import.meta.url), 'User-agent: *\nDisallow: /\n');
console.log('Vercel output ready: new console only; legacy entry points retired.');

import fs from 'node:fs';
import crypto from 'node:crypto';
const source = new URL('../../dashboard.html', import.meta.url);
const frozen = new URL('../reference/dashboard-original.html', import.meta.url);
const digest = path => crypto.createHash('sha256').update(fs.readFileSync(path)).digest('hex');
if (digest(source) !== digest(frozen)) throw new Error('Upstream dashboard changed. Review the preserved engine and parity before publishing.');
const output = new URL('../../console/', import.meta.url);
// Only this generated directory is replaced; upstream HTML and Supabase stay untouched.
fs.rmSync(output, { recursive: true, force: true });
fs.cpSync(new URL('../dist/', import.meta.url), output, { recursive: true });
fs.writeFileSync(new URL('.nojekyll', output), '');
console.log('GitHub Pages release prepared in console/.');

import fs from 'node:fs';
// Only the already-public anonymous credential belongs in the browser.
const html = fs.readFileSync(new URL('../../dashboard.html', import.meta.url), 'utf8');
const read = name => {
  const value = html.match(new RegExp(`const ${name}\\s*=\\s*['"]([^'"]+)['"]`))?.[1];
  if (!value) throw new Error(`Missing upstream ${name}`);
  return value;
};
const key = read('SUPA_KEY');
const claims = JSON.parse(Buffer.from(key.split('.')[1], 'base64url'));
if (claims.role !== 'anon') throw new Error('Browser config must never include a privileged key.');
const url = read('SUPA_URL');
if (new URL(url).hostname !== 'ftmgmfdqdqxboiktxcoj.supabase.co') throw new Error('Unexpected production project.');
fs.writeFileSync(new URL('../src/production-config.json', import.meta.url), JSON.stringify({ url, key }) + '\n');
console.log('Production config verified: existing public anonymous key.');

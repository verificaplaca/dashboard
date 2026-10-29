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
const site = { url: read('SITE_SUPA_URL'), key: read('SITE_SUPA_KEY') };
if (JSON.parse(Buffer.from(site.key.split('.')[1], 'base64url')).role !== 'anon' || new URL(site.url).hostname !== 'ozquoloetuzynnyzkado.supabase.co') throw new Error('Unexpected public checkout source.');
fs.writeFileSync(new URL('../src/production-config.json', import.meta.url), JSON.stringify({ url, key, site }) + '\n');
console.log('Production config verified: existing public anonymous key.');

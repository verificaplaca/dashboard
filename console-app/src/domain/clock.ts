export function saoPauloClock(now: Date) {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', hourCycle: 'h23' }).formatToParts(now);
  const get = (key: string) => parts.find(p => p.type === key)!.value;
  return { hour: Number(get('hour')), day: `${get('year')}-${get('month')}-${get('day')}` };
}
export function greeting(hour: number) {
  if (hour < 5) return 'Boa madrugada';
  if (hour < 12) return 'Bom dia';
  if (hour < 18) return 'Boa tarde';
  return 'Boa noite';
}

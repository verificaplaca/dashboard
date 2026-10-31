import test from 'node:test';
import assert from 'node:assert/strict';
import { chartTooltipCopy } from './chart-tooltip';

test('mobile copy preserves complete formatted money, negative values and comparison footnotes', () => {
  const value = 'CAC após transição · com aprendizado: R$ 144.066,85 de lucro total; −R$ 12.345,67 vs. manter';
  const result = chartTooltipCopy({ title: ['Após alterações', 'R$ 3.400,00/dia'], beforeBody: ['Valores estimados'], body: [{ before: ['Cenário'], lines: [value], after: ['98 dias'] }, { lines: ['Manter: R$ 0,00'] }], afterBody: ['Dia equivalente: 10/09/2026'], footer: ['Melhor resultado: manter'] });
  assert.equal(result.title, 'Após alterações · R$ 3.400,00/dia');
  assert.deepEqual(result.lines, ['Valores estimados', 'Cenário', value, '98 dias', 'Manter: R$ 0,00', 'Dia equivalente: 10/09/2026', 'Melhor resultado: manter']);
});

test('empty optional tooltip sections do not create empty rows or invent values', () => {
  assert.deepEqual(chartTooltipCopy({ body: [{ lines: ['', '   ', 'Sem dados no dia equivalente'] }] }), { title: '', lines: ['Sem dados no dia equivalente'] });
  assert.deepEqual(chartTooltipCopy({}), { title: '', lines: [] });
});

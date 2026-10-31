import test from 'node:test';
import assert from 'node:assert/strict';
import { fieldScrollDelta, visibleViewport } from './visible-viewport';

test('keyboard viewport reserves only the visible area, including Safari pan offset', () => {
  assert.deepEqual(visibleViewport(844, { height: 360, offsetTop: 80 }), { height: 360, top: 80, inset: 404 });
  assert.deepEqual(visibleViewport(844, { height: 844, offsetTop: 0 }), { height: 844, top: 0, inset: 0 });
});

test('content-resizing browsers and missing VisualViewport have no artificial keyboard gap', () => {
  assert.deepEqual(visibleViewport(320, { height: 320, offsetTop: 0 }), { height: 320, top: 0, inset: 0 });
  assert.deepEqual(visibleViewport(568), { height: 568, top: 0, inset: 0 });
  assert.deepEqual(visibleViewport(568, { height: NaN, offsetTop: -2 }), { height: 568, top: 0, inset: 0 });
});

test('focused fields scroll by the minimum needed above the action footer', () => {
  const area = { top: 80, bottom: 280 };
  assert.equal(fieldScrollDelta({ top: 260, bottom: 304 }, area), 24);
  assert.equal(fieldScrollDelta({ top: 60, bottom: 104 }, area), -20);
});

test('already visible fields stay still while keyboard animation settles', () => {
  assert.equal(fieldScrollDelta({ top: 100, bottom: 144 }, { top: 80, bottom: 280 }), 0);
  assert.equal(fieldScrollDelta({ top: 80, bottom: 124 }, { top: 80, bottom: 280 }), 0);
});

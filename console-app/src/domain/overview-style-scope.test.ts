import test from 'node:test';
import assert from 'node:assert/strict';
import {scopeOverviewThemeStyles} from './overview-style-scope';

test('legacy quoted and unquoted dark selectors address the overview host',()=>{
 for(const selector of ['[data-theme=dark]','[data-theme="dark"]',"[data-theme='dark']"]){
  assert.equal(scopeOverviewThemeStyles(`${selector} .period-reading{color:white}`),':host([data-theme="dark"]) .period-reading{color:white}');
 }
});
test('existing host selectors do not become nested or change on a second pass',()=>{
 const css=':host([data-theme="dark"]) .lit-rank{background:#222}';
 assert.equal(scopeOverviewThemeStyles(css),css);
 assert.equal(scopeOverviewThemeStyles(scopeOverviewThemeStyles(css)),css);
});
test('non-theme selectors, declarations and financial content remain unchanged',()=>{
 const css='[data-mode="summary"]{display:none}[data-theme=light]{color:#111}.chart{width:100%}';
 assert.equal(scopeOverviewThemeStyles(css),css);
});

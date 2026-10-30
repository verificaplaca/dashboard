import test from 'node:test';
import assert from 'node:assert/strict';
import {createElement,Fragment} from 'react';
import {renderToStaticMarkup} from 'react-dom/server';
import {GooeyLoader} from '../components/ui/loader-10';
test('two loader instances have independent SVG filter references',()=>{
 const html=renderToStaticMarkup(createElement(Fragment,null,createElement(GooeyLoader),createElement(GooeyLoader)));
 const ids=[...html.matchAll(/<filter id="([^"]+)"/g)].map(match=>match[1]);
 assert.equal(ids.length,2);assert.notEqual(ids[0],ids[1]);
 for(const id of ids)assert.ok(html.includes(`url(#${id})`));
});
test('loader preserves caller style, configurable colors and accessible label',()=>{
 const html=renderToStaticMarkup(createElement(GooeyLoader,{primaryColor:'#123456',secondaryColor:'#abcdef',borderColor:'#eee',className:'custom-loader',style:{width:220},'aria-label':'Aguarde'}));
 for(const value of ['--gooey-primary-color:#123456','--gooey-secondary-color:#abcdef','--gooey-border-color:#eee','width:220px','custom-loader','aria-label="Aguarde"','prefers-reduced-motion:reduce'])assert.ok(html.includes(value),value);
});

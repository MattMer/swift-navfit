// Load the actual separate ES modules used by the development browser preview.
// The packaged HTML combines module scopes and can conceal missing imports.
import assert from 'node:assert/strict';
const nodes = new Map(), handlers = new Map();
function node(id) {
  if (!nodes.has(id)) nodes.set(id, {
    innerHTML:'', textContent:'', dataset:{}, style:{},
    addEventListener(){}, querySelectorAll(){return [];},
    setAttribute(){}, focus(){}, showModal(){}, close(){}
  });
  return nodes.get(id);
}
globalThis.document = {
  querySelector:node, addEventListener(name, fn){handlers.set(name,fn);},
  body:{classList:{toggle(){},contains(){return false;}}}
};
globalThis.window = {scrollY:0,scrollTo(){},addEventListener(){}};
globalThis.location = {hash:'#overview'};
globalThis.localStorage = {getItem(){return null;},setItem(){}};
globalThis.innerWidth = 1366;
const interval = globalThis.setInterval;
globalThis.setInterval = ()=>0;
await import('../prototype/app.mjs');
globalThis.setInterval = interval;
function click(action, id, extra={}) {
  const button = {dataset:{action,id:String(id),...extra}};
  handlers.get('click')({target:{closest(selector){return selector==='[data-action]'?button:null;}},preventDefault(){}});
  assert.equal(node('#notice').textContent,'','Action must not show a runtime error');
}
function navigate(page) {
  const anchor = {hash:'#'+page};
  handlers.get('click')({target:{closest(selector){return selector==='a[href^="#"]'?anchor:null;}},preventDefault(){}});
}
click('open',1);
assert.match(node('#main').innerHTML,/Summary Group Average/);
navigate('groups');
assert.match(node('#main').innerHTML,/Open group/);
click('group-open',1);
assert.match(node('#main').innerHTML,/Shared defaults/);
click('group-tab',null,{tab:'Comparison'});
assert.match(node('#main').innerHTML,/data-comparison="Trait33"/);
assert.doesNotMatch(node('#main').innerHTML,/Freeze saved roster/);
click('goto-resources');
click('goto-seniors');
assert.match(node('#main').innerHTML,/No reporting seniors yet/);
click('senior-new');
assert.match(node('#main').innerHTML,/Previous graded trait count/);
assert.equal(node('#notice').textContent,'');
console.log('PASS: actual ES-module report and summary-group navigation');

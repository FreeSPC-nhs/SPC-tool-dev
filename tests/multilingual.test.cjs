// Run: node --test tests/localisation.test.cjs tests/multilingual.test.cjs
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
function load() {
  const entries = {};
  const memory = new Map();
  const events = {};
  const draft = {hidden: true, textContent: ''};
  const selector = {options: [], value: 'en', addEventListener: (n, cb) => events['select:'+n]=cb,
    appendChild(option) {this.options.push(option);}};
  const document = {
    readyState:'loading', documentElement:{lang:'en'},
    addEventListener(name,fn) {events[name]=fn;},
    dispatchEvent(event) {events[event.type]?.(event);},
    getElementById(id) {return id==='translationDraftNotice'?draft:id==='languageSelect'?selector:null;},
    createElement() {return {value:'',textContent:'',remove(){}};},
    querySelector(){return null;},
    querySelectorAll(){return [];}
  };
  const window = {localStorage: {getItem:k=>memory.get(k),setItem:(k,v)=>memory.set(k,v)}};
  const ctx = vm.createContext({window,document,Intl, CustomEvent:function(type,params){this.type=type;this.detail=params.detail;}});
  const exec = name => vm.runInContext(fs.readFileSync(path.join(root,name),'utf8'),ctx,{filename:name});
  exec('i18n.js');
  const register=window.SimpleSPCi18n.registerLanguage;
  window.SimpleSPCi18n.registerLanguage=(lang,dict)=>{entries[lang]=dict;register(lang,dict);};
  for(const lang of ['en','fr','es','de','it','pt']) exec('i18n/'+lang+'.js');
  events.DOMContentLoaded();
  return {i18n:window.SimpleSPCi18n,entries,document,selector,draft};
}
const languages=['en','fr','es','de','it','pt'];
test('all six languages load and offer a selector choice',()=>{
  const {i18n,selector}=load();
  assert.deepEqual(Array.from(i18n.getAvailableLanguages()),languages);
  assert.deepEqual(selector.options.map(x=>x.value),languages);
  for(const lang of languages) assert.equal(i18n.setLanguage(lang),true);
  assert.equal(i18n.setLanguage('xx'),false);
});
test('English, French, Spanish and German dictionaries are complete, and remaining draft keys are valid',()=>{
 const {entries}=load(),en=entries.en,fr=entries.fr;
 assert.equal(Object.keys(en).length,1350);
 assert.deepEqual(Object.keys(en).sort(),Object.keys(fr).sort());
 assert.deepEqual(Object.keys(en).sort(),Object.keys(entries.es).sort());
 assert.deepEqual(Object.keys(en).sort(),Object.keys(entries.de).sort());
 const tokens=value=>Array.from(String(value).matchAll(/\{([a-zA-Z0-9_]+)\}/g),x=>x[1]).sort();
 for(const lang of ['es','de','it','pt']) {
  assert.ok(Object.keys(entries[lang]).length>=400,lang);
  for(const [k,value] of Object.entries(entries[lang])) {
    assert.ok(Object.hasOwn(en,k),lang+': '+k);
    assert.deepEqual(tokens(value),tokens(en[k]),lang+': '+k);
  }
 }
});
test('number parsing uses original data language, not UI language',()=>{
 const {i18n}=load();
 for(const lang of ['de','es','it','pt']) {
  i18n.setLanguage(lang);
  assert.equal(i18n.parseNumber('1.234'),1234,lang+' dot grouping');
  assert.equal(i18n.parseNumber('1,234'),1.234,lang+' comma decimal');
  assert.equal(i18n.parseNumber('1.234,56'),1234.56,lang+' combined');
  assert.equal(i18n.parseNumber('1,234.56'),1234.56,lang+' English combined');
  assert.equal(i18n.parseNumber('1\u00a0234,56'),1234.56,lang+' NBSP');
  assert.equal(i18n.parseNumber('-12,5 %'),-12.5,lang+' percentage points');
  assert.equal(i18n.parseNumber('1.234',{language:'en'}),1.234,lang+' imported English');
 }
 i18n.setLanguage('en'); assert.equal(i18n.parseNumber('1,234'),1234);
 i18n.setLanguage('fr'); assert.equal(i18n.parseNumber('1.234'),1.234);
 for(const value of ['1,2,3','1.23.45','12 34','foo','2026-01-01']) assert.ok(Number.isNaN(i18n.parseNumber(value)),value);
});
test('localized display and percent formatting',()=>{
 const {i18n}=load();
 for(const [lang,locale] of [['en','en-GB'],['fr','fr-FR'],['es','es-ES'],['de','de-DE'],['it','it-IT'],['pt','pt-PT']]) {
  i18n.setLanguage(lang);
  assert.equal(i18n.getLocale(),locale);
  assert.equal(i18n.formatFixed(12.5,2),lang==='en'?'12.50':'12,50');
  assert.equal(i18n.formatPercent(12.5,1),new Intl.NumberFormat(locale,{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1,useGrouping:false}).format(0.125));
 }
});
test('draft notice is visible only for unreviewed languages',()=>{
 const {i18n,draft,document}=load();
 for(const lang of ['it','pt']) {
  i18n.setLanguage(lang);
  assert.equal(draft.hidden,false);
  assert.ok(draft.textContent.length>20);
  assert.equal(document.documentElement.lang,lang);
  assert.notEqual(i18n.translate('buttons.generate'), 'Generate / Recalculate');
  assert.equal(i18n.translate('helper.response.empty'), 'Please type a question about SPC or your chart.');
 }
 for(const lang of ['en','fr','es','de']) {i18n.setLanguage(lang);assert.equal(draft.hidden,true);}
 i18n.setLanguage('es');
 assert.notEqual(i18n.translate('helper.response.empty'), 'Please type a question about SPC or your chart.');
});

test('Spanish wizard has no English fallback text except chart-name symbols',()=>{
 const {i18n,entries}=load();
 i18n.setLanguage('es');
 const unchangedAllowed=new Set(['wizard.chart.xmr','wizard.chart.xbars','wizard.chart.t','wizard.chart.g']);
 for(const k of Object.keys(entries.en).filter(k=>k.startsWith('wizard.'))) {
   assert.ok(Object.hasOwn(entries.es,k),'missing Spanish wizard entry: '+k);
   if(!unchangedAllowed.has(k)) {
     assert.notEqual(entries.es[k],entries.en[k],'English wizard fallback: '+k);
   }
 }
});


test('German dictionary is complete and contains no English fallback for explanatory text',()=>{
 const {i18n,entries,draft}=load();
 i18n.setLanguage('de');
 assert.equal(draft.hidden,true);
 const allowed=new Set(['auto.axis.s','visual.xbar','wizard.chart.xmr','wizard.chart.xbars','wizard.chart.t','wizard.chart.g']);
 for(const k of Object.keys(entries.en)) {
  assert.ok(Object.hasOwn(entries.de,k),k);
  assert.ok(String(entries.de[k]).trim(),k);
  if(!allowed.has(k) && entries.en[k].trim().length>30) {
   assert.notEqual(entries.de[k].trim(),entries.en[k].trim(),'English text: '+k);
  }
 }
 assert.match(i18n.translate('helper.faq.common-vs-special'),/gewöhnliche|besonder/);
 assert.match(i18n.translate('wizard.reason.paretoRows'),/Pareto/);
 assert.equal(i18n.formatFixed(12.5,2),'12,50');
});

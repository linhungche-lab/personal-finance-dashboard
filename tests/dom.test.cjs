// DOM/state tests without native browser access. Optional development-only linkedom/postcss.
const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {parseHTML} = require('linkedom');
const postcss = require('postcss');
const {XLSX,workbook,bytes} = require('./fixtures.cjs');
const ExcelImport = require('../excel-import.js');
const root=path.join(__dirname,'..');
const html=fs.readFileSync(path.join(root,'index.html'),'utf8');
const app=fs.readFileSync(path.join(root,'app.js'),'utf8');
function setup(protocol='file:') {
  const {document}=parseHTML(html);
  const instances=[];
  class Chart {
    constructor(canvas, config){this.canvas=canvas;this.config=config;instances.push(this);}
    destroy(){this.destroyed=true;}
  }
  class Worker {
    constructor() {
      this.self={postMessage:data=>queueMicrotask(()=>{if(!this.stopped)this.onmessage({data});})};
      const context=vm.createContext({self:this.self,globalThis:null,Uint8Array,ArrayBuffer,TextDecoder,TextEncoder});
      context.globalThis=context;
      context.importScripts=(...files)=>files.forEach(file=>vm.runInContext(fs.readFileSync(path.join(root,file),'utf8'),context));
      vm.runInContext(fs.readFileSync(path.join(root,'excel-worker.js'),'utf8'),context);
      this.context=context;
    }
    postMessage(data){queueMicrotask(()=>this.self.onmessage({data}));}
    terminate(){this.stopped=true;}
  }
  const context=vm.createContext({document,location:{protocol},ExcelImport,XLSX,Chart,Worker,Date,Promise,setTimeout,clearTimeout,console});
  vm.runInContext(app,context);
  const get=id=>document.getElementById(id);
  function file(wb,name='synthetic.xlsx') {
    const buffer=bytes(wb);
    return {name,size:buffer.length,arrayBuffer:async()=>buffer.buffer.slice(buffer.byteOffset,buffer.byteOffset+buffer.length)};
  }
  return {context,get,document,instances,file};
}

test('HTML IDs, static scripts, app element references, CSS syntax and mobile status rule',()=>{
  const {document}=parseHTML(html);
  const ids=[...document.querySelectorAll('[id]')].map(el=>el.id);
  assert.equal(new Set(ids).size,ids.length);
  for(const [,id] of app.matchAll(/getElementById\("([^"]+)"\)/g))assert.ok(document.getElementById(id),id);
  for(const script of document.querySelectorAll('script[src]')) {
    const src=script.getAttribute('src');
    assert.equal(/^https?:/.test(src),false);
    assert.ok(fs.existsSync(path.join(root,src)));
  }
  const css=fs.readFileSync(path.join(root,'style.css'),'utf8');
  assert.ok(postcss.parse(css).nodes.length>0);
  assert.equal(/\.data-status\s*\{\s*display:\s*none/.test(css),false);
  assert.match(document.querySelector('meta[http-equiv="Content-Security-Policy"]').content,/connect-src 'none'/);
});
for(const protocol of ['file:','https:'])test(`${protocol} imports, error recovery, UI replacement and safe workbook text`,async()=>{
  const {context,get,document,instances,file}=setup(protocol);
  assert.equal(get('dataStatusLabel').textContent,'尚未載入財務資料');
  assert.ok(get('dashboardContent').classList.contains('hidden'));
  const wb=workbook();wb.Sheets['持股明細'].C7={t:'s',v:'<img src=x onerror=alert(1)>'};
  await context.importExcelFile(file(wb,'<script>test</script>.xlsx'));
  assert.equal(get('dataStatusLabel').textContent,'✓ 財務資料已載入',get('importErrorList').textContent);
  assert.equal(get('loadedFileName').textContent,'<script>test</script>.xlsx');
  assert.equal(get('loadedFileName').children.length,0);
  assert.equal(get('portfolioMarketValue').textContent,'NT$ 410');
  assert.equal(document.querySelectorAll('#holdingsBody tr').length,2);
  assert.equal(document.querySelectorAll('#holdingsBody img').length,0);
  assert.equal(instances[0].config.data.datasets[0].data.length,5);
  assert.match(get('loadedTime').textContent,/載入時間/);
  const oldTime=get('loadedTime').textContent;
  const invalid=workbook();delete invalid.Sheets['財務資料輸入'].E5;
  await context.importExcelFile(file(invalid,'bad.xlsx'));
  assert.equal(get('importErrors').classList.contains('hidden'),false);
  assert.match(get('importErrorList').textContent,/E5/);
  assert.match(get('importErrorContext').textContent,/上一份/);
  assert.equal(get('portfolioMarketValue').textContent,'NT$ 410');
  assert.equal(get('loadedTime').textContent,oldTime);
  assert.equal(get('loadedFileName').textContent,'<script>test</script>.xlsx');
  await context.importExcelFile(file(workbook({empty:true}),'empty.xlsx'));
  assert.equal(get('portfolioMarketValue').textContent,'NT$ 0');
  assert.equal(get('largestHoldingName').textContent,'目前無持股');
  assert.equal(get('largestHoldingPortfolioWeight').textContent,'—');
  assert.equal(document.querySelectorAll('#holdingsBody tr').length,0);
  assert.equal(get('importErrors').classList.contains('hidden'),true);
  assert.equal(get('chooseExcelBtn').disabled,false);
  assert.equal(get('excelFileInput').value,'');
  assert.ok(instances.slice(0,3).every(chart=>chart.destroyed));
  const fresh=setup(protocol);
  assert.ok(fresh.get('dashboardContent').classList.contains('hidden'));
  assert.equal(fresh.get('loadedFileName').textContent,'');
});
test('first invalid import stays unloaded and can recover with the same filename',async()=>{
  const {context,get,file}=setup();
  const wb=workbook();wb.Sheets['持股明細'].H7={t:'n',f:'1+1',v:2};
  await context.importExcelFile(file(wb));
  assert.ok(get('dashboardContent').classList.contains('hidden'));
  assert.equal(get('chooseExcelBtn').textContent,'選擇 Excel');
  assert.match(get('importErrorList').textContent,/不能使用公式/);
  await context.importExcelFile(file(workbook()));
  assert.equal(get('dashboardContent').classList.contains('hidden'),false);
});
test('last selection wins when asynchronous file reads arrive out of order',async()=>{
  const {context,get,file}=setup();
  let release;
  const older=file(workbook(),'old.xlsx');const read=older.arrayBuffer;
  older.arrayBuffer=()=>new Promise(resolve=>{release=async()=>resolve(await read());});
  const first=context.importExcelFile(older);
  await context.importExcelFile(file(workbook({empty:true}),'new.xlsx'));
  await release();await first;
  assert.equal(get('loadedFileName').textContent,'new.xlsx');
  assert.equal(get('portfolioMarketValue').textContent,'NT$ 0');
});
test('missing parser and rejected File read report errors without unhandled exceptions',async()=>{
  const {context,get,file}=setup();
  await context.importExcelFile({name:'read.xlsx',size:50,arrayBuffer:async()=>{throw new Error('denied');}});
  assert.equal(get('importErrors').classList.contains('hidden'),false);
  assert.equal(get('chooseExcelBtn').disabled,false);
  context.ExcelImport=undefined;
  await context.importExcelFile(file(workbook()));
  assert.equal(get('importErrors').classList.contains('hidden'),false);
});

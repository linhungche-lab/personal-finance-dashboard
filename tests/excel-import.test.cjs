const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const {XLSX, workbook, bytes} = require('./fixtures.cjs');
const A = require('../excel-import.js');
const context = vm.createContext({document:{addEventListener(){}}, console});
vm.runInContext(fs.readFileSync(path.join(__dirname,'../app.js'),'utf8'), context);
const parse = wb => A.parseBuffer(bytes(wb), XLSX);
function fails(wb, pattern) {
  assert.throws(()=>parse(wb), error=>error instanceof A.ImportError && error.issues.some(issue=>pattern.test(issue)));
}

test('xlsx maps inputs and original model independently recomputes all key totals',()=>{
  const raw = parse(workbook());
  assert.equal(Object.hasOwn(raw.assets, 'stocks'), false);
  assert.equal(raw.holdings[0].ticker, '0050');
  const {model} = context.buildFinancialModel(raw);
  assert.equal(model.assets.stocks, 410);
  assert.equal(model.totalAssets, 5410);
  assert.equal(model.totalLiabilities, 5000);
  assert.equal(model.netWorth, 410);
  assert.equal(model.portfolioCost, 220);
  assert.equal(model.portfolioPnl, 190);
  assert.equal(model.portfolioReturn, 190/220);
  assert.equal(model.metrics.emergencyFundMonths, 24);
  assert.equal(model.holdings[0].ticker, 'TEST');
});
test('legacy xls uses the same validated model',()=>{
  assert.equal(A.parseBuffer(bytes(workbook(),'biff8'),XLSX).holdings.length,2);
});
test('does not access spec sheet or depend on calculated cache values',()=>{
  const wb = workbook();
  Object.defineProperty(wb.Sheets,'V1財務模型規格',{get(){throw new Error('must not access spec');}});
  const ws = wb.Sheets['持股明細'];
  ws.K7={t:'n',f:'F7*H7*I7',v:999999999};
  ws['!ref']='A1:K8';
  const raw = A.workbookToRaw(wb,XLSX);
  assert.equal(Object.hasOwn(raw.holdings[0],'marketValueTwd'),false);
  assert.equal(context.buildFinancialModel(raw).model.assets.stocks,410);
});
test('parser requests only the two permitted sheets',()=>{
  let options;
  const wrapper={...XLSX,read(data,opts){options=opts;return XLSX.read(data,opts);}};
  A.parseBuffer(bytes(workbook()),wrapper);
  assert.deepEqual(options.sheets,['財務資料輸入','持股明細']);
  assert.equal(options.cellFormula,true);
});
test('input formulas are rejected even with a valid numeric cache',()=>{
  for(const [sheet,cell] of [['財務資料輸入','E5'],['持股明細','H7']]) {
    const wb=workbook();wb.Sheets[sheet][cell]={t:'n',f:'1+1',v:2};fails(wb,/不能使用公式/);
  }
});
test('blank holding template rows and calculated-only rows do not become positions',()=>{
  const wb=workbook({empty:true}); const ws=wb.Sheets['持股明細'];
  ws.K7={t:'n',f:'0',v:0}; ws['!ref']='A1:K8';
  assert.deepEqual(parse(wb).holdings,[]);
});
test('zero balances are distinct from missing balances',()=>{
  const wb=workbook(); wb.Sheets['財務資料輸入'].E5={t:'n',v:0};
  assert.equal(parse(wb).assets.realEstate,0);
  delete wb.Sheets['財務資料輸入'].E5; fails(wb,/E5.*未填寫/);
});
test('negative, malformed, boolean, date, percentage and error inputs are rejected',()=>{
  for(const cell of [{t:'n',v:-1},{t:'s',v:'12,34'},{t:'s',v:'$100'},{t:'b',v:true},{t:'d',v:new Date()},{t:'e',v:7},{t:'n',v:.1,z:'0%'}]) {
    const wb=workbook(); wb.Sheets['財務資料輸入'].E5=cell; fails(wb,/E5/);
  }
});
test('well-formed thousands-separated numeric text is accepted',()=>{
  const wb=workbook(); wb.Sheets['財務資料輸入'].E5={t:'s',v:'1,234.50'};
  assert.equal(parse(wb).assets.realEstate,1234.5);
});
test('missing sheets, headers, field keys and duplicate fields are reported',()=>{
  const wb=workbook(); delete wb.Sheets['持股明細']; wb.SheetNames=wb.SheetNames.filter(x=>x!=='持股明細');fails(wb,/缺少工作表/);
  const h=workbook();h.Sheets['持股明細'].F6={t:'s',v:'Wrong'};fails(h,/完整表頭/);
  const f=workbook();f.Sheets['財務資料輸入'].C5={t:'s',v:'unknown'};fails(f,/未支援/);
  const d=workbook();d.Sheets['財務資料輸入'].C6={t:'s',v:'realEstate'};fails(d,/重複/);
});
test('partial holdings, missing costs, invalid FX and wrong units cannot understate stocks',()=>{
  for(const [sheet,ref,cell] of [
    ['持股明細','B7',null],['持股明細','G7',null],['持股明細','F7',{t:'n',v:0}],
    ['持股明細','I8',{t:'n',v:0}],['持股明細','I7',{t:'n',v:30}],
    ['持股明細','E7',{t:'s',v:'<img>'}],['財務資料輸入','D5',{t:'s',v:'USD'}]
  ]) { const wb=workbook();if(cell)wb.Sheets[sheet][ref]=cell;else delete wb.Sheets[sheet][ref];fails(wb,/持股明細|單位/); }
});
test('reads holdings past original template row 106',()=>{
  const wb=workbook();XLSX.utils.sheet_add_aoa(wb.Sheets['持股明細'],[['台股','EXTRA','','','TWD',1,1,2,1]],{origin:'A107'});
  assert.equal(parse(wb).holdings.length,3);
});
test('headers tolerate spaces/newlines and changed column order',()=>{
  const wb=workbook();const ws=wb.Sheets['持股明細'];
  for(let r=6;r<=8;r++){[ws['F'+r],ws['I'+r]]=[ws['I'+r],ws['F'+r]];}
  assert.equal(parse(wb).holdings[1].shares,2);
});
test('file validation and unreadable payloads produce actionable errors',()=>{
  for(const file of [{name:'a.csv',size:10},{name:'a.xlsx',size:0},{name:'a.xlsx',size:A.MAX_BYTES+1}])assert.throws(()=>A.validateFile(file),A.ImportError);
  assert.throws(()=>A.parseBuffer(Buffer.from('invalid'),XLSX),A.ImportError);
  assert.throws(()=>A.parseBuffer(Buffer.from('PKbroken'),XLSX),A.ImportError);
  assert.throws(()=>A.parseBuffer(bytes(workbook()),null),A.ImportError);
});
test('huge ranges and unsafe aggregate amounts are rejected',()=>{
  const wb=workbook();wb.Sheets['持股明細']['!fullref']='A1:O10001';
  assert.throws(()=>A.workbookToRaw(wb,XLSX),/Excel/);
  const large=workbook();large.Sheets['財務資料輸入'].E5={t:'n',v:Number.MAX_SAFE_INTEGER};fails(large,/合計/);
});
test('provided template schema accepts synthetic inputs without modifying its source',()=>{
  const file=path.join(__dirname,'../sample/Personal_Finance_Dashboard_V1.1_Template.xlsx');
  if(!fs.existsSync(file))return; // sample files are intentionally not deployed
  const input=fs.readFileSync(file);
  const wb=XLSX.read(input,{sheets:['財務資料輸入','持股明細'],cellFormula:true});
  const finance=wb.Sheets['財務資料輸入'];
  for(let row=5;row<=23;row++)if(finance['C'+row]?.v!=='stocks')finance['E'+row]={t:'n',v:0};
  // Only synthetic data in memory; discard any pre-existing holding input rows.
  const holdings=wb.Sheets['持股明細'];
  const range=XLSX.utils.decode_range(holdings['!ref']);
  for(let r=6;r<=range.e.r;r++)for(let c=0;c<9;c++)delete holdings[XLSX.utils.encode_cell({r,c})];
  const raw=A.workbookToRaw(wb,XLSX);
  assert.equal(raw.holdings.length,0);
  assert.equal(context.buildFinancialModel(raw).model.totalAssets,0);
  assert.equal(fs.readFileSync(file).equals(input),true);
});

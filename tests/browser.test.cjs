// Optional end-to-end test: requires Playwright + its headless Chromium runtime.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const http = require('node:http');
const {pathToFileURL} = require('node:url');
const {chromium} = require('playwright');
const {workbook, bytes} = require('./fixtures.cjs');
const root = path.join(__dirname,'..');
const allowed = new Set(['index.html','style.css','app.js','excel-import.js','excel-worker.js','vendor/chart-4.4.7.umd.min.js','vendor/xlsx-0.20.3.full.min.js']);
const server = http.createServer((req,res)=>{
  const name = req.url === '/' ? 'index.html' : req.url.slice(1);
  if(req.method !== 'GET' || !allowed.has(name)){res.writeHead(404);res.end();return;}
  res.setHeader('Content-Type',name.endsWith('.js')?'text/javascript':name.endsWith('.css')?'text/css':'text/html');
  res.end(fs.readFileSync(path.join(root,name)));
});
async function main() {
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const base=`http://127.0.0.1:${server.address().port}/`;
  const browser=await chromium.launch({headless:true});
  try {
    const context=await browser.newContext({viewport:{width:1440,height:1000}});
    const page=await context.newPage();
    const errors=[];const requests=[];
    page.on('pageerror',error=>errors.push(error.message));
    context.on('request',req=>requests.push({url:req.url(),method:req.method()}));
    await context.addInitScript(()=>{
      // Any attempt to persist financial data fails the test immediately.
      Storage.prototype.setItem=function(){throw new Error('Storage writes forbidden');};
      indexedDB.open=function(){throw new Error('IndexedDB writes forbidden');};
    });
    await page.goto(base);
    assert.equal(await page.locator('#dataStatusLabel').textContent(),'尚未載入財務資料');
    assert.equal(await page.locator('#dashboardContent').isVisible(),false);
    async function upload(wb,name='synthetic.xlsx') {
      await page.locator('#excelFileInput').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:bytes(wb)});
      await page.waitForFunction(()=>!document.getElementById('chooseExcelBtn').disabled);
    }
    const valid=workbook();
    valid.Sheets['持股明細'].C7={t:'s',v:'<img src=x onerror=alert(1)>'};
    await upload(valid);
    assert.equal(await page.locator('#dataStatusLabel').textContent(),'✓ 財務資料已載入');
    assert.equal(await page.locator('#portfolioMarketValue').textContent(),'NT$ 410');
    assert.equal(await page.locator('#holdingsBody tr').count(),2);
    assert.equal(await page.locator('#holdingsBody img').count(),0);
    assert.match(await page.locator('#loadedTime').textContent(),/載入時間/);
    assert.equal(await page.locator('#excelFileInput').inputValue(),'');
    assert.equal(await page.evaluate(()=>Chart.getChart('assetChart')?.data.datasets[0].data.length),5);
    await page.screenshot({path:path.join(root,'work/desktop.png'),fullPage:true});
    const invalid=workbook();delete invalid.Sheets['財務資料輸入'].E5;
    await upload(invalid,'bad.xlsx');
    assert.equal(await page.locator('#importErrors').isVisible(),true);
    assert.match(await page.locator('#importErrorList').textContent(),/E5/);
    assert.equal(await page.locator('#portfolioMarketValue').textContent(),'NT$ 410');
    assert.equal(await page.locator('#loadedFileName').textContent(),'synthetic.xlsx');
    assert.match(await page.locator('#importErrorContext').textContent(),/上一份/);
    await upload(workbook({empty:true}),'empty.xlsx');
    assert.equal(await page.locator('#portfolioMarketValue').textContent(),'NT$ 0');
    assert.equal(await page.locator('#largestHoldingName').textContent(),'目前無持股');
    assert.equal(await page.locator('#holdingsBody tr').count(),0);
    assert.equal(await page.locator('#importErrors').isVisible(),false);
    await page.setViewportSize({width:375,height:812});
    await upload(workbook(),'a-very-long-synthetic-finance-filename-for-mobile.xlsx');
    assert.equal(await page.locator('#dataStatus').isVisible(),true);
    assert.equal(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),true);
    await page.screenshot({path:path.join(root,'work/mobile.png'),fullPage:true});
    await page.reload();
    assert.equal(await page.locator('#dashboardContent').isVisible(),false);
    assert.equal(await page.locator('#loadedFileName').textContent(),'');
    await page.locator('#excelFileInput').setInputFiles({name:'broken.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from('bad')});
    await page.waitForFunction(()=>!document.getElementById('chooseExcelBtn').disabled);
    assert.equal(await page.locator('#importErrors').isVisible(),true);
    assert.equal(await page.locator('#dashboardContent').isVisible(),false);
    await upload(workbook()); // recover from first-load failure
    assert.equal(await page.locator('#dashboardContent').isVisible(),true);
    assert.deepEqual(errors,[]);
    assert.equal(requests.every(r=>r.url.startsWith(base)&&r.method==='GET'),true);
    // Direct double-click remains supported, with local main-thread fallback.
    const local=await context.newPage();
    local.on('pageerror',error=>errors.push(error.message));
    await local.goto(pathToFileURL(path.join(root,'index.html')).href);
    await local.locator('#excelFileInput').setInputFiles({name:'local.xlsx',mimeType:'application/octet-stream',buffer:bytes(workbook())});
    await local.waitForFunction(()=>!document.getElementById('chooseExcelBtn').disabled);
    assert.equal(await local.locator('#portfolioMarketValue').textContent(),'NT$ 410');
    assert.deepEqual(errors,[]);
    console.log('PASS: browser import, workers, recalculation, reload, error recovery, stale-data prevention, escaping, mobile layout, no storage writes or outgoing data requests, and file:// fallback.');
  } finally { await browser.close(); }
}
main().catch(error=>{console.error(error);process.exitCode=1;}).finally(()=>server.close());

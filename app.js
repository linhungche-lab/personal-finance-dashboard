const COLORS = {
  primary:"#142D44", blue:"#58768C", teal:"#648C83", gold:"#B7975E",
  gray:"#9CA4AA", success:"#34735A", warning:"#A87827", danger:"#A64A45", neutral:"#77828B"
};

let charts={asset:null,liability:null,portfolio:null};
let loadedRaw = null;
let importSequence = 0;

document.addEventListener("DOMContentLoaded",()=>{
  const input = document.getElementById("excelFileInput");
  document.getElementById("chooseExcelBtn").addEventListener("click",()=>input.click());
  input.addEventListener("change",()=>{
    const file = input.files[0];
    if (file) importExcelFile(file);
  });
});

function parseExcelLocally(buffer) {
  // Workers keep parsing off the UI thread on GitHub Pages / HTTP(S).
  // file:// disallows workers in some browsers; use the same local parser there.
  if (location.protocol === "file:" || typeof Worker === "undefined") {
    return Promise.resolve().then(()=>ExcelImport.parseBuffer(buffer, globalThis.XLSX));
  }
  return new Promise((resolve, reject)=>{
    const worker = new Worker("excel-worker.js");
    const timer = setTimeout(()=>finish(null, ["Excel 解析超過 15 秒，請縮小檔案後重試。"]), 15000);
    function finish(raw, issues) {
      clearTimeout(timer);
      worker.terminate();
      if (issues) reject(new ExcelImport.ImportError(issues));
      else resolve(raw);
    }
    worker.onmessage = ({data})=>finish(data.raw, data.issues);
    worker.onerror = event=>{
      event.preventDefault();
      finish(null, ["Excel 解析器無法啟動，請確認 excel-worker.js 與 vendor 資料夾完整。"]);
    };
    worker.postMessage(buffer, [buffer]);
  });
}

async function importExcelFile(file) {
  const sequence = ++importSequence;
  const button = document.getElementById("chooseExcelBtn");
  const input = document.getElementById("excelFileInput");
  button.disabled = true;
  input.disabled = true;
  button.textContent = "載入中…";
  document.getElementById("importControls").setAttribute("aria-busy", "true");
  showImportIssues([]);
  try {
    if (!globalThis.ExcelImport) throw new Error("parser unavailable");
    ExcelImport.validateFile(file);
    const buffer = await file.arrayBuffer();
    if (sequence !== importSequence) return;
    const raw = await parseExcelLocally(buffer);
    if (sequence !== importSequence) return;
    const loadedAt = new Date();
    raw.meta = {label: file.name, loadedAt: loadedAt.toISOString()};
    // The adapter validates the complete workbook before the existing model is called.
    document.getElementById("dashboardContent").classList.remove("hidden");
    renderDashboard(raw);
    loadedRaw = raw;
    document.getElementById("unloadedState").classList.add("hidden");
    setText("dataStatusLabel", "✓ 財務資料已載入");
    setText("loadedFileName", file.name);
    setText("loadedTime", `載入時間：${loadedAt.toLocaleString("zh-TW", {hour12:false})}`);
    document.getElementById("dataStatus").classList.add("is-loaded");
  } catch (error) {
    if (sequence !== importSequence) return;
    if (!loadedRaw) document.getElementById("dashboardContent").classList.add("hidden");
    else renderDashboard(loadedRaw);
    showImportIssues(error.issues || ["無法完成 Excel 載入，請確認檔案與本機解析套件完整後再試。"], file.name);
  } finally {
    if (sequence === importSequence) {
      input.value = ""; // Allow selecting the same file again after correcting it.
      input.disabled = false;
      button.disabled = false;
      button.textContent = loadedRaw ? "重新選擇 Excel" : "選擇 Excel";
      document.getElementById("importControls").setAttribute("aria-busy", "false");
    }
  }
}

function showImportIssues(issues, filename = "") {
  const panel = document.getElementById("importErrors");
  const list = document.getElementById("importErrorList");
  list.replaceChildren();
  panel.classList.toggle("hidden", !issues.length);
  if (!issues.length) return;
  setText("importErrorTitle", `未匯入：${filename}`);
  setText("importErrorContext", loadedRaw ? "目前仍顯示上一份成功載入的資料。請修正以下問題後重新選擇 Excel。" : "尚未載入任何資料。請修正以下問題後重新選擇 Excel。");
  for (const issue of issues) {
    const li = document.createElement("li");
    li.textContent = issue;
    list.appendChild(li);
  }
  panel.focus();
}

function renderDashboard(raw){
  const {model,warnings}=buildFinancialModel(raw);
  setText("snapshotLabel",raw.meta?.label||"本機 Excel");
  renderOverview(model);
  renderAssetAllocation(model);
  renderLiabilityStructure(model);
  renderFinancialHealth(model);
  renderObservations(model);
  renderPortfolio(model);
  renderHoldingsTable(model.holdings);
  if(typeof Chart === "undefined") warnings.push("圖表套件未載入；數值仍可查看，請確認 vendor 資料夾完整。");
  renderWarnings(warnings);
}

function buildFinancialModel(raw){
  const warnings=[];
  const holdings=normalizeHoldings(raw.holdings||[],warnings);
  const stocks=sum(holdings.map(h=>h.marketValueTwd));

  const assets={
    realEstate:positiveNumber(raw.assets?.realEstate),
    stocks,
    cash:positiveNumber(raw.assets?.cash),
    foreignCash:positiveNumber(raw.assets?.foreignCash),
    otherInvestments:positiveNumber(raw.assets?.otherInvestments),
    otherAssets:positiveNumber(raw.assets?.otherAssets)
  };

  const liabilities={
    mortgage:positiveNumber(raw.liabilities?.mortgage),
    personalLoan:positiveNumber(raw.liabilities?.personalLoan),
    investmentLoan:positiveNumber(raw.liabilities?.investmentLoan),
    otherLoans:positiveNumber(raw.liabilities?.otherLoans),
    otherLiabilities:positiveNumber(raw.liabilities?.otherLiabilities)
  };

  const income={
    salaryIncome:positiveNumber(raw.income?.salaryIncome),
    dividendIncome:positiveNumber(raw.income?.dividendIncome),
    rentalIncome:positiveNumber(raw.income?.rentalIncome),
    otherIncome:positiveNumber(raw.income?.otherIncome)
  };

  const expenses={
    livingExpenses:positiveNumber(raw.expenses?.livingExpenses),
    mortgageInterest:positiveNumber(raw.expenses?.mortgageInterest),
    otherInterest:positiveNumber(raw.expenses?.otherInterest),
    fixedExpenses:positiveNumber(raw.expenses?.fixedExpenses)
  };

  const totalAssets=sum(Object.values(assets));
  const totalLiabilities=sum(Object.values(liabilities));
  const netWorth=totalAssets-totalLiabilities;
  const totalIncome=sum(Object.values(income));
  const totalCash=assets.cash+assets.foreignCash;
  const liquidAssets=totalCash+assets.stocks;
  const passiveIncome=income.dividendIncome+income.rentalIncome;

  const metrics={
    debtRatio:safeDivide(totalLiabilities,totalAssets),
    liquidAssetRatio:safeDivide(liquidAssets,totalAssets),
    mortgageLTV:(liabilities.mortgage===0&&assets.realEstate===0)?null:safeDivide(liabilities.mortgage,assets.realEstate),
    debtIncomeMultiple:totalIncome===0?(totalLiabilities===0?0:null):totalLiabilities/totalIncome,
    cashRatio:safeDivide(totalCash,totalAssets),
    realEstateConcentration:safeDivide(assets.realEstate,totalAssets),
    stockConcentration:safeDivide(assets.stocks,totalAssets),
    passiveIncomeRatio:safeDivide(passiveIncome,totalIncome),
    emergencyFundMonths:expenses.livingExpenses===0?null:totalCash/(expenses.livingExpenses/12)
  };

  if(liabilities.mortgage>0&&assets.realEstate===0) warnings.push("房貸餘額存在，但不動產市值為 0，無法計算房貸 LTV。");
  if(totalLiabilities>0&&totalIncome===0) warnings.push("存在負債，但年收入為 0，無法計算負債收入倍數。");
  if(expenses.livingExpenses===0) warnings.push("年生活支出為 0，無法計算緊急預備金月數。");

  const portfolioCost=sum(holdings.map(h=>h.costBasisTwd||0));
  const portfolioPnl=stocks-portfolioCost;
  const portfolioReturn=portfolioCost>0?portfolioPnl/portfolioCost:null;

  const holdingsWithWeights=holdings.map(h=>({
    ...h,
    portfolioWeight:stocks>0?h.marketValueTwd/stocks:null,
    totalAssetWeight:totalAssets>0?h.marketValueTwd/totalAssets:null
  })).sort((a,b)=>b.marketValueTwd-a.marketValueTwd);

  return {model:{
    assets,liabilities,income,expenses,holdings:holdingsWithWeights,
    totalAssets,totalLiabilities,netWorth,totalIncome,totalCash,liquidAssets,passiveIncome,
    metrics,portfolioCost,portfolioPnl,portfolioReturn
  },warnings};
}

function normalizeHoldings(rows,warnings){
  return rows.map((item,index)=>{
    const ticker=String(item.ticker||"").trim();
    const shares=positiveNumber(item.shares);
    const price=positiveNumber(item.currentPrice);
    const fx=positiveNumber(item.fxToTwd);
    const cost=(item.avgCost===null||item.avgCost==="")?null:positiveNumber(item.avgCost);

    if(!ticker&&shares===0&&price===0) return null;
    if(!ticker) warnings.push(`持股第 ${index+1} 筆缺少股票代號。`);
    if(shares<=0) warnings.push(`${ticker||`持股第 ${index+1} 筆`} 缺少有效股數。`);
    if(price<=0) warnings.push(`${ticker||`持股第 ${index+1} 筆`} 缺少有效目前價格。`);
    if(fx<=0) warnings.push(`${ticker||`持股第 ${index+1} 筆`} 缺少有效匯率。`);

    const marketValueLocal=shares*price;
    const marketValueTwd=marketValueLocal*fx;
    const costBasisTwd=cost==null?null:shares*cost*fx;
    const unrealizedPnlTwd=costBasisTwd==null?null:marketValueTwd-costBasisTwd;
    const unrealizedReturn=costBasisTwd>0?unrealizedPnlTwd/costBasisTwd:null;

    return {...item,ticker,shares,currentPrice:price,fxToTwd:fx,avgCost:cost,
      marketValueLocal,marketValueTwd,costBasisTwd,unrealizedPnlTwd,unrealizedReturn};
  }).filter(Boolean);
}

function renderOverview(m){
  setText("totalAssetsShort",formatMoneyShort(m.totalAssets));
  setText("totalAssetsFull",formatMoneyFull(m.totalAssets));
  setText("totalLiabilitiesShort",formatMoneyShort(m.totalLiabilities));
  setText("totalLiabilitiesFull",formatMoneyFull(m.totalLiabilities));
  setText("netWorthShort",formatMoneyShort(m.netWorth));
  setText("netWorthFull",formatMoneyFull(m.netWorth));
  setText("debtRatioValue",formatPercent(m.metrics.debtRatio));
  const s=getDebtRatioStatus(m.metrics.debtRatio);
  const el=document.getElementById("debtRatioStatus");
  el.className=`status-line status-${s.tone}`;
  el.textContent=`● ${s.label}`;
}

function renderAssetAllocation(m){
  const items=[
    {label:"不動產",value:m.assets.realEstate,color:COLORS.primary},
    {label:"股票 / ETF",value:m.assets.stocks,color:COLORS.blue},
    {label:"現金",value:m.totalCash,color:COLORS.teal},
    {label:"其他投資",value:m.assets.otherInvestments,color:COLORS.gold},
    {label:"其他資產",value:m.assets.otherAssets,color:COLORS.gray}
  ].filter(x=>x.value>0);
  setText("assetAllocationTotal",formatMoneyFull(m.totalAssets));
  renderLegend("assetLegend",items,m.totalAssets);
  charts.asset=replaceDoughnut(charts.asset,"assetChart",items,"總資產",formatMoneyShort(m.totalAssets));
}

function renderLiabilityStructure(m){
  const items=[
    {label:"房貸",value:m.liabilities.mortgage,color:COLORS.primary},
    {label:"信貸",value:m.liabilities.personalLoan,color:COLORS.blue},
    {label:"投資借款",value:m.liabilities.investmentLoan,color:COLORS.gold},
    {label:"其他貸款",value:m.liabilities.otherLoans,color:COLORS.teal},
    {label:"其他負債",value:m.liabilities.otherLiabilities,color:COLORS.gray}
  ].filter(x=>x.value>0);

  setText("liabilityStructureTotal",formatMoneyFull(m.totalLiabilities));
  const content=document.getElementById("liabilityChartContent");
  const empty=document.getElementById("noLiabilityState");

  if(m.totalLiabilities===0){
    content.classList.add("hidden"); empty.classList.remove("hidden");
    if(charts.liability){charts.liability.destroy();charts.liability=null}
    return;
  }
  content.classList.remove("hidden"); empty.classList.add("hidden");
  renderLegend("liabilityLegend",items,m.totalLiabilities);
  charts.liability=replaceDoughnut(charts.liability,"liabilityChart",items,"總負債",formatMoneyShort(m.totalLiabilities));
}

function renderFinancialHealth(m){
  const items=[
    ["負債比",formatPercent(m.metrics.debtRatio),getDebtRatioStatus(m.metrics.debtRatio)],
    ["房貸 LTV",formatPercent(m.metrics.mortgageLTV),getMortgageLtvStatus(m.metrics.mortgageLTV)],
    ["流動資產比例",formatPercent(m.metrics.liquidAssetRatio),getLiquidityStatus(m.metrics.liquidAssetRatio)],
    ["負債收入倍數",formatMultiple(m.metrics.debtIncomeMultiple),getDebtIncomeStatus(m.metrics.debtIncomeMultiple)],
    ["緊急預備金",formatMonths(m.metrics.emergencyFundMonths),getEmergencyFundStatus(m.metrics.emergencyFundMonths)]
  ];
  const grid=document.getElementById("healthGrid");
  grid.innerHTML="";
  items.forEach(([label,value,status])=>{
    const card=document.createElement("article");
    card.className="health-card";
    card.style.setProperty("--status-color",status.color);
    card.innerHTML=`<span class="health-label">${label}</span><strong class="health-value">${value}</strong><span class="health-status">${status.label}</span>`;
    grid.appendChild(card);
  });
}

function renderObservations(m){
  const items=[
    ["不動產集中度",formatPercent(m.metrics.realEstateConcentration),getRealEstateStatus(m.metrics.realEstateConcentration)],
    ["股票集中度",formatPercent(m.metrics.stockConcentration),getStockStatus(m.metrics.stockConcentration)],
    ["現金比例",formatPercent(m.metrics.cashRatio),getCashStatus(m.metrics.cashRatio)],
    ["被動收入比例",formatPercent(m.metrics.passiveIncomeRatio),getPassiveIncomeStatus(m.metrics.passiveIncomeRatio)]
  ];
  const list=document.getElementById("observationList");
  list.innerHTML="";
  items.forEach(([name,value,status])=>{
    const row=document.createElement("div");
    row.className="observation-row";
    row.innerHTML=`<span class="observation-name">${name}</span><strong class="observation-value">${value}</strong><span class="observation-status" style="color:${status.color}">${status.label}</span>`;
    list.appendChild(row);
  });
}

function renderPortfolio(m){
  setText("portfolioMarketValue",formatMoneyFull(m.assets.stocks));
  setText("portfolioCost",formatMoneyFull(m.portfolioCost));
  setText("portfolioPnl",formatSignedMoney(m.portfolioPnl));
  setText("portfolioReturn",formatSignedPercent(m.portfolioReturn));
  setPnlClass("portfolioPnl",m.portfolioPnl);
  setPnlClass("portfolioReturn",m.portfolioReturn);

  const largest=m.holdings[0];
  if(largest){
    setText("largestHoldingName",`${largest.ticker}${largest.name?` · ${largest.name}`:""}`);
    setText("largestHoldingPortfolioWeight",`占股票投資組合 ${formatPercent(largest.portfolioWeight)}`);
    setText("largestHoldingAssetWeight",`占總資產 ${formatPercent(largest.totalAssetWeight)}`);
  }

  else {
    setText("largestHoldingName","目前無持股");
    setText("largestHoldingPortfolioWeight","—");
    setText("largestHoldingAssetWeight","—");
  }

  const palette=["#142D44","#31556F","#58768C","#648C83","#B7975E","#8C7D70","#9CA4AA"];
  const items=m.holdings.map((h,i)=>({label:h.ticker,value:h.marketValueTwd,color:palette[i%palette.length]}));
  charts.portfolio=replaceDoughnut(charts.portfolio,"portfolioChart",items,"股票 / ETF",formatMoneyShort(m.assets.stocks));
}

function renderHoldingsTable(rows){
  const tbody=document.getElementById("holdingsBody");
  tbody.innerHTML="";
  rows.forEach(h=>{
    const tr=document.createElement("tr");
    const pnlClass=getPnlClass(h.unrealizedPnlTwd);
    const retClass=getPnlClass(h.unrealizedReturn);
    tr.innerHTML=`
      <td><span class="holding-symbol">${escapeHtml(h.ticker||"—")}</span>${h.name?`<span class="holding-name">${escapeHtml(h.name)}</span>`:""}</td>
      <td>${escapeHtml(h.market||"—")}</td>
      <td class="num">${formatShares(h.shares)}</td>
      <td class="num">${escapeHtml(formatPrice(h.currentPrice,h.currency))}</td>
      <td class="num">${formatMoneyFull(h.marketValueTwd)}</td>
      <td class="num">${formatPercent(h.portfolioWeight)}</td>
      <td class="num ${pnlClass}">${formatSignedMoney(h.unrealizedPnlTwd)}</td>
      <td class="num ${retClass}">${formatSignedPercent(h.unrealizedReturn)}</td>`;
    tbody.appendChild(tr);
  });
}

function renderWarnings(warnings){
  const section=document.getElementById("validationSection");
  const list=document.getElementById("validationList");
  if(!warnings.length){section.classList.add("hidden");list.innerHTML="";return}
  list.innerHTML=warnings.map(w=>`<li>${escapeHtml(w)}</li>`).join("");
  section.classList.remove("hidden");
}

function replaceDoughnut(existing,id,items,label,value){
  if(existing) existing.destroy();
  if(typeof Chart==="undefined") return null;
  const canvas=document.getElementById(id);
  const centerPlugin={
    id:`center-${id}`,
    afterDraw(chart){
      const {ctx,chartArea}=chart;if(!chartArea)return;
      const x=(chartArea.left+chartArea.right)/2,y=(chartArea.top+chartArea.bottom)/2;
      ctx.save();ctx.textAlign="center";ctx.textBaseline="middle";
      ctx.fillStyle=COLORS.neutral;ctx.font='600 12px Calibri, "Segoe UI", sans-serif';ctx.fillText(label,x,y-10);
      ctx.fillStyle=COLORS.primary;ctx.font='700 18px Calibri, "Segoe UI", sans-serif';ctx.fillText(value,x,y+14);ctx.restore();
    }
  };
  return new Chart(canvas,{
    type:"doughnut",
    data:{labels:items.map(x=>x.label),datasets:[{data:items.map(x=>x.value),backgroundColor:items.map(x=>x.color),borderColor:"#fff",borderWidth:3,hoverOffset:4}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:"70%",plugins:{legend:{display:false},tooltip:{callbacks:{label(ctx){const total=sum(ctx.dataset.data);const v=Number(ctx.raw||0);return ` ${ctx.label}: ${formatMoneyFull(v)} (${formatPercent(total>0?v/total:null)})`;}}}}},
    plugins:[centerPlugin]
  });
}

function renderLegend(id,items,total){
  const c=document.getElementById(id);c.innerHTML="";
  items.forEach(item=>{
    const row=document.createElement("div");row.className="legend-row";
    row.innerHTML=`<span class="legend-swatch" style="background:${item.color}"></span><span class="legend-label">${item.label}</span><strong class="legend-value">${formatPercent(total>0?item.value/total:null)}</strong>`;
    c.appendChild(row);
  });
}

function status(label,tone){return {label,tone,color:{success:COLORS.success,warning:COLORS.warning,danger:COLORS.danger,neutral:COLORS.neutral}[tone]}}
function getDebtRatioStatus(v){if(v==null)return status("無資料","neutral");if(v<.3)return status("健康","success");if(v<=.5)return status("注意","warning");return status("高負債","danger")}
function getLiquidityStatus(v){if(v==null)return status("無資料","neutral");if(v>=.3)return status("流動性充足","success");if(v>=.15)return status("一般","warning");return status("流動性偏低","danger")}
function getMortgageLtvStatus(v){if(v==null)return status("不適用 / 無資料","neutral");if(v<.5)return status("健康","success");if(v<=.7)return status("注意","warning");return status("高槓桿","danger")}
function getDebtIncomeStatus(v){if(v==null)return status("無年收入資料","neutral");if(v<3)return status("健康","success");if(v<=6)return status("注意","warning");return status("槓桿偏高","danger")}
function getEmergencyFundStatus(v){if(v==null)return status("無資料","neutral");if(v>=12)return status("非常充足","success");if(v>=6)return status("健康","success");if(v>=3)return status("注意","warning");return status("偏低","danger")}
function getRealEstateStatus(v){if(v==null)return status("無資料","neutral");if(v<.5)return status("一般","neutral");if(v<=.7)return status("不動產配置偏高","warning");return status("不動產高度集中","warning")}
function getStockStatus(v){if(v==null)return status("無資料","neutral");if(v<.4)return status("一般","neutral");if(v<=.6)return status("股票配置較高","warning");return status("股票配置高度集中","warning")}
function getCashStatus(v){if(v==null)return status("無資料","neutral");if(v<.05)return status("現金水位偏低","warning");if(v<.1)return status("現金水位較低","warning");if(v<=.25)return status("一般","neutral");if(v<=.4)return status("現金配置偏高","warning");return status("現金配置明顯偏高","warning")}
function getPassiveIncomeStatus(v){if(v==null)return status("無資料","neutral");if(v<.2)return status("主要依賴主動收入","neutral");if(v<=.5)return status("被動收入占比提升","success");return status("被動收入占比較高","success")}

function formatMoneyFull(v){if(!Number.isFinite(v))return "—";return `${v<0?"-":""}NT$ ${Math.round(Math.abs(v)).toLocaleString("zh-TW")}`}
function formatMoneyShort(v){if(!Number.isFinite(v))return "—";const s=v<0?"-":"";const a=Math.abs(v);if(a>=1e8)return `${s}NT$ ${(a/1e8).toFixed(2)}億`;if(a>=1e4)return `${s}NT$ ${(a/1e4).toFixed(1)}萬`;return `${s}NT$ ${Math.round(a).toLocaleString("zh-TW")}`}
function formatSignedMoney(v){if(!Number.isFinite(v))return "—";if(v===0)return "NT$ 0";return `${v>0?"+":"-"}NT$ ${Math.round(Math.abs(v)).toLocaleString("zh-TW")}`}
function formatPercent(v){return Number.isFinite(v)?`${(v*100).toFixed(1)}%`:"—"}
function formatSignedPercent(v){if(!Number.isFinite(v))return "—";if(v===0)return "0.0%";return `${v>0?"+":"-"}${Math.abs(v*100).toFixed(1)}%`}
function formatMultiple(v){return Number.isFinite(v)?`${v.toFixed(1)}x`:"—"}
function formatMonths(v){return Number.isFinite(v)?`${v.toFixed(1)} 月`:"—"}
function formatShares(v){return Number.isFinite(v)?v.toLocaleString("zh-TW",{maximumFractionDigits:4}):"—"}
function formatPrice(v,c){return Number.isFinite(v)?`${c||""} ${v.toLocaleString("zh-TW",{maximumFractionDigits:2})}`.trim():"—"}
function safeDivide(a,b){return Number.isFinite(a)&&Number.isFinite(b)&&b!==0?a/b:null}
function positiveNumber(v){const n=Number(v);return Number.isFinite(n)&&n>=0?n:0}
function sum(a){return a.reduce((t,v)=>t+(Number(v)||0),0)}
function setText(id,v){const e=document.getElementById(id);if(e)e.textContent=v}
function getPnlClass(v){if(!Number.isFinite(v)||v===0)return "pnl-neutral";return v>0?"pnl-positive":"pnl-negative"}
function setPnlClass(id,v){const e=document.getElementById(id);e.classList.remove("pnl-positive","pnl-negative","pnl-neutral");e.classList.add(getPnlClass(v))}
function escapeHtml(v){return String(v??"").replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;")}

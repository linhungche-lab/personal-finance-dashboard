// Entirely synthetic inputs; never embed data from a user's workbook.
const XLSX = require('../vendor/xlsx-0.20.3.full.min.js');
const groups = {
  assets: ['realEstate','cash','foreignCash','otherInvestments','otherAssets'],
  liabilities: ['mortgage','personalLoan','investmentLoan','otherLoans','otherLiabilities'],
  income: ['salaryIncome','dividendIncome','rentalIncome','otherIncome'],
  expenses: ['livingExpenses','mortgageInterest','otherInterest','fixedExpenses']
};
const header = ['市場','股票代號','名稱','類型','幣別','持有股數','平均成本/股','目前價格/股','匯率\n(TWD/1原幣)','市值\n(原幣)','市值\n(TWD)'];
function workbook({empty = false} = {}) {
  const rows = [['Synthetic test'], [], [], ['分類','欄位名稱','程式欄位','單位','輸入值 / 自動值']];
  for (const [group, keys] of Object.entries(groups)) {
    for (const key of keys) rows.push([group,key,key,['income','expenses'].includes(group)?'TWD / 年':'TWD',1000]);
  }
  rows.push(['assets','ignored total','stocks','TWD',{t:'n',f:'1+1',v:999999}]);
  const holdings = [['Synthetic test'],[],[],[],[],header];
  if (!empty) holdings.push(['台股','0050','Test ETF','ETF','TWD',10,4,5,1],['美股','TEST','Test Inc','股票','USD',2,3,6,30]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '財務資料輸入');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(holdings), '持股明細');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Ignored spec']]), 'V1財務模型規格');
  return wb;
}
function bytes(wb, bookType = 'xlsx') { return XLSX.write(wb, {type:'buffer',bookType}); }
module.exports = {XLSX, workbook, bytes};

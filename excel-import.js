/* Local Excel adapter. No storage, network calls, or Excel writes. */
(function (root) {
  "use strict";
  const SHEETS = ["財務資料輸入", "持股明細"];
  const MAX_BYTES = 5 * 1024 * 1024;
  const MAX_ROWS = 10000;
  const MAX_COLUMNS = 64;
  const FIELDS = {
    assets: ["realEstate", "cash", "foreignCash", "otherInvestments", "otherAssets"],
    liabilities: ["mortgage", "personalLoan", "investmentLoan", "otherLoans", "otherLiabilities"],
    income: ["salaryIncome", "dividendIncome", "rentalIncome", "otherIncome"],
    expenses: ["livingExpenses", "mortgageInterest", "otherInterest", "fixedExpenses"]
  };
  const HOLDING_COLUMNS = {
    market: "市場", ticker: "股票代號", name: "名稱", type: "類型", currency: "幣別",
    shares: "持有股數", avgCost: "平均成本/股", currentPrice: "目前價格/股", fxToTwd: "匯率(TWD/1原幣)"
  };
  const normalized = value => String(value ?? "").normalize("NFKC").replace(/\s+/g, "").trim();
  const hasFormula = cell => cell && (cell.f !== undefined || cell.F !== undefined);
  const isBlank = cell => !hasFormula(cell) && (!cell || cell.v == null || String(cell.v).trim() === "");

  class ImportError extends Error {
    constructor(issues) { super("Excel 資料格式需要修正"); this.name = "ImportError"; this.issues = issues; }
  }

  function validateFile(file) {
    const issues = [];
    if (!/\.(xlsx|xls)$/i.test(file.name)) issues.push("請選擇 .xlsx 或 .xls Excel 檔案。");
    if (!file.size) issues.push("檔案是空的，請重新選擇。");
    if (file.size > MAX_BYTES) issues.push("檔案超過 5 MB，請移除不需要的圖片或資料後再試。");
    if (issues.length) throw new ImportError(issues);
  }

  function parseBuffer(buffer, XLSX) {
    if (!XLSX || typeof XLSX.read !== "function") throw new ImportError(["Excel 解析套件未載入，請確認 vendor 資料夾完整後重新整理。"]);
    if (!buffer.byteLength || buffer.byteLength > MAX_BYTES) throw new ImportError(["Excel 檔案必須介於 1 byte 與 5 MB 之間。"]);
    const bytes = new Uint8Array(buffer);
    const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    const ole = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((v, i) => bytes[i] === v);
    if (!zip && !ole) throw new ImportError(["檔案不是支援的 Excel 活頁簿；請用 Excel 另存成 .xlsx 後再試。"]);
    let workbook;
    try {
      workbook = XLSX.read(buffer, {
        type: "array", sheets: SHEETS, cellFormula: true, cellText: true,
        cellDates: true, sheetRows: MAX_ROWS + 1, bookVBA: false, cellHTML: false
      });
    } catch (_) {
      throw new ImportError(["無法解析 Excel。檔案可能已損毀、受密碼保護或格式不受支援，請另存成未加密 .xlsx 後再試。"]);
    }
    return workbookToRaw(workbook, XLSX);
  }

  function workbookToRaw(workbook, XLSX) {
    const issues = [];
    const raw = { assets: {}, liabilities: {}, income: {}, expenses: {}, holdings: [] };
    const address = (r, c) => XLSX.utils.encode_cell({ r, c });
    function sheetInfo(name) {
      const sheet = workbook.Sheets[name];
      if (!sheet) { issues.push(`缺少工作表「${name}」。`); return null; }
      if (!sheet["!ref"]) { issues.push(`「${name}」是空白工作表。`); return null; }
      const range = XLSX.utils.decode_range(sheet["!fullref"] || sheet["!ref"]);
      if (range.e.r >= MAX_ROWS || range.e.c >= MAX_COLUMNS) {
        issues.push(`「${name}」超過 ${MAX_ROWS} 列或 ${MAX_COLUMNS} 欄的匯入上限，請移除多餘範圍。`);
        return null;
      }
      return { sheet, range, name };
    }
    function findHeader(info, required) {
      for (let r = info.range.s.r; r <= Math.min(info.range.e.r, 29); r++) {
        const columns = {};
        for (let c = info.range.s.c; c <= info.range.e.c; c++) {
          const cell = info.sheet[address(r, c)];
          if (isBlank(cell) || hasFormula(cell)) continue;
          const key = normalized(cell.v);
          if (required.some(label => normalized(label) === key)) {
            if (columns[key] !== undefined) columns[key] = -1;
            else columns[key] = c;
          }
        }
        if (required.every(label => columns[normalized(label)] !== undefined)) {
          for (const label of required) if (columns[normalized(label)] < 0) issues.push(`「${info.name}」第 ${r + 1} 列的「${label}」表頭重複。`);
          return { row: r, columns };
        }
      }
      issues.push(`「${info.name}」前 30 列找不到完整表頭，需要：${required.join("、")}。`);
      return null;
    }
    function cellValue(info, r, c, label, { text = false, optional = false, positive = false } = {}) {
      const ref = address(r, c);
      const cell = info.sheet[ref];
      const where = `「${info.name}」${ref}（${label}）`;
      if (hasFormula(cell)) { issues.push(`${where} 必須填原始輸入值，不能使用公式或其快取值。`); return null; }
      if (isBlank(cell)) {
        if (!optional) issues.push(`${where} 未填寫${text ? "。" : "，無金額請明確填 0。"}`);
        return text ? "" : null;
      }
      if (cell.t === "e" || cell.t === "b" || cell.t === "d") {
        issues.push(`${where} 不是有效的${text ? "文字" : "數字"}。`); return null;
      }
      if (text) {
        // Preserve explicit Excel zero-padding for numeric ticker cells (e.g. 0050).
        const value = String(cell.t === "n" && cell.w ? cell.w : cell.v).trim();
        if (value.length > 120) issues.push(`${where} 不得超過 120 字元。`);
        return value;
      }
      const source = typeof cell.v === "number" ? cell.v : String(cell.v).trim();
      const numericText = /^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/;
      const number = typeof source === "number" ? source : numericText.test(source) ? Number(source.replaceAll(",", "")) : NaN;
      if (!Number.isFinite(number) || number < 0 || number > Number.MAX_SAFE_INTEGER || (positive && number === 0)) {
        issues.push(`${where} 必須是${positive ? "大於 0" : "大於或等於 0"}的有限數字（不可含貨幣符號或百分比）。`);
        return null;
      }
      if (typeof cell.w === "string" && cell.w.includes("%")) {
        issues.push(`${where} 不可使用百分比格式。`); return null;
      }
      return number;
    }

    const finance = sheetInfo(SHEETS[0]);
    if (finance) {
      const header = findHeader(finance, ["程式欄位", "輸入值 / 自動值", "單位"]);
      if (header && Object.values(header.columns).every(c => c >= 0)) {
        const keyColumn = header.columns[normalized("程式欄位")];
        const valueColumn = header.columns[normalized("輸入值 / 自動值")];
        const unitColumn = header.columns[normalized("單位")];
        const groups = new Map(Object.entries(FIELDS).flatMap(([group, fields]) => fields.map(key => [key, group])));
        const seen = new Set();
        for (let r = header.row + 1; r <= finance.range.e.r; r++) {
          if (isBlank(finance.sheet[address(r, keyColumn)])) {
            if (!isBlank(finance.sheet[address(r, valueColumn)])) issues.push(`「${finance.name}」第 ${r + 1} 列有數值但缺少程式欄位。`);
            continue;
          }
          const key = cellValue(finance, r, keyColumn, "程式欄位", { text: true });
          if (key === "stocks") continue; // Never read the Excel stock total, including its formula cache.
          if (!groups.has(key)) { issues.push(`「${finance.name}」第 ${r + 1} 列包含未支援的程式欄位，請依 V1.1 模板填寫。`); continue; }
          if (seen.has(key)) { issues.push(`「${finance.name}」第 ${r + 1} 列的 ${key} 重複。`); continue; }
          seen.add(key);
          const group = groups.get(key);
          const unit = cellValue(finance, r, unitColumn, "單位", { text: true });
          const expectedUnit = group === "income" || group === "expenses" ? "TWD/年" : "TWD";
          if (normalized(unit) !== expectedUnit) issues.push(`「${finance.name}」第 ${r + 1} 列 ${key} 的單位須為 ${expectedUnit}，請勿填月金額或原幣金額。`);
          raw[group][key] = cellValue(finance, r, valueColumn, key);
        }
        for (const key of groups.keys()) if (!seen.has(key)) issues.push(`「${finance.name}」缺少程式欄位 ${key}。`);
      }
    }

    const holdings = sheetInfo(SHEETS[1]);
    if (holdings) {
      const header = findHeader(holdings, Object.values(HOLDING_COLUMNS));
      if (header && Object.values(header.columns).every(c => c >= 0)) {
        const cols = Object.fromEntries(Object.entries(HOLDING_COLUMNS).map(([key, label]) => [key, header.columns[normalized(label)]]));
        for (let r = header.row + 1; r <= holdings.range.e.r; r++) {
          // Blank template rows have calculated J:O cells; only inspect original input columns.
          if (Object.values(cols).every(c => isBlank(holdings.sheet[address(r, c)]))) continue;
          // The template ends with a merged instructional footer, not a holding record.
          const footer = (holdings.sheet["!merges"] || []).some(m => m.s.r === r && m.e.r === r && m.s.c <= cols.market && m.e.c >= cols.fxToTwd);
          if (footer && Object.values(cols).filter(c => c !== cols.market).every(c => isBlank(holdings.sheet[address(r, c)]))) continue;
          const item = {};
          for (const [key, c] of Object.entries(cols)) {
            const text = ["market", "ticker", "name", "type", "currency"].includes(key);
            item[key] = cellValue(holdings, r, c, HOLDING_COLUMNS[key], {
              text, optional: key === "name" || key === "type", positive: ["shares", "currentPrice", "fxToTwd"].includes(key)
            });
          }
          item.currency = String(item.currency || "").toUpperCase();
          if (!/^[A-Z]{3}$/.test(item.currency)) issues.push(`「${holdings.name}」第 ${r + 1} 列的幣別須為三碼，例如 TWD 或 USD。`);
          if (item.currency === "TWD" && item.fxToTwd !== 1) issues.push(`「${holdings.name}」第 ${r + 1} 列的 TWD 匯率必須是 1。`);
          const marketValue = item.shares * item.currentPrice * item.fxToTwd;
          const costValue = item.shares * item.avgCost * item.fxToTwd;
          if (![marketValue, costValue].every(v => Number.isFinite(v) && v <= Number.MAX_SAFE_INTEGER)) issues.push(`「${holdings.name}」第 ${r + 1} 列的市值或成本超出可計算範圍。`);
          raw.holdings.push(item);
        }
      }
    }
    if (issues.length) throw new ImportError(issues);
    // Keep aggregate arithmetic within JavaScript's safely representable range.
    const amounts = Object.keys(FIELDS).flatMap(group => Object.values(raw[group]));
    amounts.push(...raw.holdings.map(h => h.shares * Math.max(h.currentPrice, h.avgCost) * h.fxToTwd));
    if (amounts.reduce((a, b) => a + b, 0) > Number.MAX_SAFE_INTEGER) throw new ImportError(["資料合計超出可安全計算的範圍，請檢查金額、股數與匯率。"]);
    return raw;
  }
  const api = { parseBuffer, workbookToRaw, validateFile, ImportError, MAX_BYTES };
  root.ExcelImport = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(globalThis);

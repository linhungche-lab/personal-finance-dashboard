/* Runs only against bytes explicitly selected by the user. */
importScripts("vendor/xlsx-0.20.3.full.min.js", "excel-import.js");
self.onmessage = ({ data }) => {
  try {
    self.postMessage({ raw: ExcelImport.parseBuffer(data, XLSX) });
  } catch (error) {
    self.postMessage({ issues: error.issues || ["Excel 解析失敗，請確認檔案格式後再試。"] });
  }
};

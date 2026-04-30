/**
 * ╔══════════════════════════════════════════════════════════╗
 * ║  刮傷要因發掘表 — Google Apps Script 擴充片段            ║
 * ║                                                          ║
 * ║  使用方式：                                               ║
 * ║  1. 開啟現有 observation-system 的 GAS 專案              ║
 * ║  2. 將下方 doPost 的 case 區塊貼入你原有的 doPost switch  ║
 * ║  3. 將 saveScratchRecord / getScratchRecords 兩個函式     ║
 * ║     貼到 GAS 檔案底部                                    ║
 * ║  4. 重新部署為「新版本」的 Web App                       ║
 * ╚══════════════════════════════════════════════════════════╝
 */

/* ─────────────────────────────────────────────
   STEP 1  貼入你原有 doPost 的 switch-case 內
   ───────────────────────────────────────────── */

// 在你的 doPost(e) 函式裡，找到 switch(body.action) 的地方，
// 加入以下兩個 case：

/*
  case "saveScratchRecord":
    result = saveScratchRecord(body.record);
    break;

  case "getScratchRecords":
    result = getScratchRecords();
    break;

  case "updateScratchIssue":
    result = updateScratchIssue(body.recordId, body.checkId, body.field, body.value);
    break;
*/


/* ─────────────────────────────────────────────
   STEP 2  貼到 GAS 檔案底部（新增兩個函式）
   ───────────────────────────────────────────── */

const SCRATCH_SHEET_NAME = "刮傷巡檢";

// 查檢項目 ID 清單（與前端 CHECK_ITEMS 順序一致）
const SCRATCH_ITEM_IDS = [
  "s01","s02","s03","s04","s05","s06","s07",
  "s08","s09","s10","s11","s12",
  "s13","s14","s15","s16",
  "s17",
  "s18","s19","s20"
];

/**
 * 儲存一筆刮傷巡檢紀錄
 * @param {Object} record - 來自前端的表單物件
 */
function saveScratchRecord(record) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SCRATCH_SHEET_NAME);

    // 若工作表不存在，自動建立並寫入標頭列
    if (!sheet) {
      sheet = ss.insertSheet(SCRATCH_SHEET_NAME);
      const headers = [
        "ID", "送出時間", "實施日", "工站名稱", "作業者", "觀察者", "觀察車型", "勤務區分",
        // 每個查檢項目三欄：判定 / 問題點 / 對策日
        ...SCRATCH_ITEM_IDS.flatMap(id => [`${id}_判定`, `${id}_問題點`, `${id}_對策日`]),
        "備注Memo"
      ];
      sheet.appendRow(headers);
      // 凍結標頭列
      sheet.setFrozenRows(1);
    }

    // 建立資料列
    const checks = record.checks || {};
    const row = [
      record.id            || Date.now(),
      record.submittedAt   || new Date().toISOString(),
      record.date          || "",
      record.station       || "",
      record.operator      || "",
      record.observer      || "",
      record.carModel      || "",
      record.shift         || "",
      // 展開每個查檢項目的三個欄位
      ...SCRATCH_ITEM_IDS.flatMap(id => {
        const c = checks[id] || {};
        return [c.result || "O", c.issue || "", c.actionDate || ""];
      }),
      record.memo || ""
    ];

    sheet.appendRow(row);
    return { status: "ok", id: record.id };

  } catch (e) {
    Logger.log("saveScratchRecord error: " + e.toString());
    return { status: "error", message: e.toString() };
  }
}

/**
 * 更新單筆刮傷巡檢紀錄中某個查檢項目的改善欄位
 * @param {string} recordId  - 紀錄 ID（即 A 欄的值）
 * @param {string} checkId   - 查檢項目 ID，例如 "s01"
 * @param {string} field     - 欄位名稱："status" | "improveNote" | "actionDate" | "issue"
 * @param {string} value     - 要寫入的值
 */
function updateScratchIssue(recordId, checkId, field, value) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SCRATCH_SHEET_NAME);
    if (!sheet) return { status: "error", message: "找不到工作表" };

    const da
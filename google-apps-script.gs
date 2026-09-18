/**
 * 楊梅高中交通安全學習網 — Google Apps Script Web App
 * ------------------------------------------------------------
 * 用途：接收學習網送出的「班級、座號、姓名、答對題數、日期時間」，
 *       寫入 Google 試算表的「記錄」工作表。
 *
 * 部署步驟：
 * 1. 開啟一份 Google 試算表。
 * 2. 選單「擴充功能 → Apps Script」，把本檔內容全部貼進去。
 * 3. 點選單「部署 → 新增部署作業 → 類型選『網頁應用程式』」。
 *    - 執行身分：我
 *    - 具有存取權的使用者：任何人
 * 4. 部署後會得到一個網址（.../exec 結尾），複製它。
 * 5. 把該網址貼到學習網 HTML 中的 GAS_URL 常數（取代 XXXXXXX）。
 *
 * 注意：網頁以 fetch(no-cors) 送出 text/plain 的 JSON 字串，
 *       因此用 e.postData.contents 讀取。
 */

var SHEET_NAME = '記錄';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }
    // 若工作表是空的，先加上標題列
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['填答日期時間', '班級', '座號', '姓名', '答對題數', '總題數']);
    }

    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var now = data.datetime || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      now,
      data.klass || '',
      data.seat || '',
      data.name || '',
      data.correct != null ? data.correct : '',
      data.total != null ? data.total : ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 供瀏覽器直接開啟測試用
function doGet() {
  return ContentService
    .createTextOutput('楊梅高中交通安全學習網 Web App 運作中。請以 POST 送出資料。')
    .setMimeType(ContentService.MimeType.TEXT);
}

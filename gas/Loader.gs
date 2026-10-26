/**
 * MPScore-Daily - 學生平時成績歷程系統 GitHub 自動熱更新引導核心
 * 網站作者 / 系統規劃：楊家驊老師 (Chiahua Yang)
 *
 * 只要在 Apps Script 貼上這段程式碼，未來 GitHub (chy23/MPScore-Daily) 有任何演算法或功能更新，
 * 您的 Google 試算表都會全自動即時抓取最新版，免去重複手動貼代碼的麻煩！
 * 
 * 部署指引：
 * 1. 貼上後儲存。
 * 2. 第一次使用請先點上方「執行」> 選擇「setup」並執行，完成 Google 帳號授權。
 * 3. 點選右上角「部署」>「新部署」，選「網頁應用程式」，【誰可以存取】務必選【所有人 (Anyone)】！
 */
var GITHUB_CORE_URL = "https://raw.githubusercontent.com/chy23/MPScore-Daily/main/gas/Code.gs";
var _cachedEngine = null;

function setup() {
  getEngine();
  SpreadsheetApp.getActiveSpreadsheet();
}

function getEngine(forceReload) {
  if (!_cachedEngine || forceReload) {
    var t = new Date().getTime();
    var response = UrlFetchApp.fetch(GITHUB_CORE_URL + "?t=" + t);
    var code = response.getContentText();
    (1, eval)(code);
    if (typeof MPScoreDailyEngine !== 'undefined') {
      _cachedEngine = MPScoreDailyEngine;
    } else {
      throw new Error("無法載入 MPScore-Daily 核心模組，請檢查 GitHub 網路連線。");
    }
  }
  return _cachedEngine;
}

// 靜態建立選單（不使用 UrlFetchApp，確保試算表每次開啟時都能 100% 穩定出現選單！）
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MPScore-Daily 平時成績工具')
    .addItem('🛠️ 建立初始表格（空白試算表第一次使用）', 'setupStandardSheets')
    .addSeparator()
    .addItem('➕ 新增平時成績科目分頁（自動帶入學生名單）', 'createNewSubjectSheetFromRoster')
    .addItem('📤 將「學生名冊」姓名同步至各科平時成績表', 'syncRosterToSubjectSheets')
    .addItem('📥 從目前成績分頁匯入名單至「學生名冊」', 'syncStudentsFromSheet')
    .addItem('🔑 為全班產生「家長查詢代碼」', 'generateTokensForStudents')
    .addSeparator()
    .addItem('🛡️ 設定/查看「教師安全金鑰 (PIN)」', 'manageTeacherAdminKeyMenu')
    .addItem('📱 查看【家長閱覽紀錄】日誌分頁', 'openAccessLogsSheet')
    .addItem('🗑️ 清空全班【家長閱覽紀錄】與次數歸零', 'clearAllAccessLogsMenu')
    .addItem('📊 查看全班各科平時總均與分項平均統計', 'previewClassStatistics')
    .addItem('📐 自動整理欄寬與表格排版', 'autoFormatAndResizeAllSheets')
    .addItem('🚀 立即更新雲端資料（改完分數點此）', 'clearSystemCache')
    .addSeparator()
    .addItem('📶 檢查系統連線狀態', 'testGitHubConnection')
    .addToUi();
}

// 選單動作分發 (由 GitHub 核心即時執行最新邏輯)
function setupStandardSheets() { getEngine().setupStandardSheets(); }
function createNewSubjectSheetFromRoster() { getEngine().createNewSubjectSheetFromRoster(); }
function syncRosterToSubjectSheets() { getEngine().syncRosterToSubjectSheets(false); }
function syncStudentsFromSheet() { getEngine().syncStudentsFromSheet(); }
function generateTokensForStudents() { getEngine().generateTokensForStudents(); }
function manageTeacherAdminKeyMenu() { getEngine().manageTeacherAdminKeyMenu(); }
function openAccessLogsSheet() { getEngine().openAccessLogsSheet(); }
function clearAllAccessLogsMenu() { getEngine().clearAllAccessLogsMenu(); }
function previewClassStatistics() { getEngine().previewClassStatistics(); }
function autoFormatAndResizeAllSheets() { getEngine().autoFormatAndResizeAllSheets(); }
function clearSystemCache() {
  _cachedEngine = null;
  getEngine(true).clearSystemCache();
}

// Web API 入口點 (附完整錯誤捕捉保護，絕不回傳 HTML 當機頁面)
function doGet(e) {
  try {
    var force = (e && e.parameter && (e.parameter.reload === '1' || e.parameter.nocache === '1'));
    return getEngine(force).handleDoGet(e);
  } catch (err) {
    var out = {
      status: "error",
      success: false,
      message: "【Loader 熱更新執行異常】: " + err.toString() + "\n\n💡 快速解決：\n1. 請在 Apps Script 上方將函式選為「setup」，並點擊「執行」完成帳號授權。"
    };
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  return doGet(e);
}

function testGitHubConnection() {
  try {
    var engine = getEngine(true);
    SpreadsheetApp.getUi().alert("🎉 連線成功！已成功自 GitHub (chy23/MPScore-Daily) 載入最新版 MPScore-Daily 平時成績歷程核心模組。");
  } catch (err) {
    SpreadsheetApp.getUi().alert("⚠️ 連線失敗: " + err.toString());
  }
}

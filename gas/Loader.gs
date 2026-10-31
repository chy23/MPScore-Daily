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
    .addSeparator()
    .addItem('📝 查看【成績異動日誌】分頁', 'openScoreUpdateLogsSheet')
    .addItem('🗑️ 清空【成績異動日誌】', 'clearScoreUpdateLogsMenu')
    .addSeparator()
    .addItem('⚙️ 設定 School LINE Hub 中樞連線', 'setupLinePushSettingsMenu')
    .addItem('📢 發送【平時成績已更新】LINE 通知給已綁定家長', 'sendLineScoreUpdateNotificationMenu')
    .addSeparator()
    .addItem('📊 查看全班各科平時總均與分項平均統計', 'previewClassStatistics')
    .addItem('➕ 新增 1 欄評量（自動建立今日日期與下拉選單）', 'addNewAssessmentColumnMenu')
    .addItem('📐 自動整理欄寬與表格排版（方案 C 版面）', 'autoFormatAndResizeAllSheets')
    .addItem('🔄 將現有成績分頁轉換為「方案 C 版面」', 'convertAllSubjectSheetsToSchemeC')
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
function openScoreUpdateLogsSheet() { getEngine().openScoreUpdateLogsSheet(); }
function clearScoreUpdateLogsMenu() { getEngine().clearScoreUpdateLogsMenu(); }
function convertAllSubjectSheetsToSchemeC() { getEngine().convertAllSubjectSheetsToSchemeC(); }
function setupLinePushSettingsMenu() { getEngine().setupLinePushSettingsMenu(); }
function sendLineScoreUpdateNotificationMenu() { getEngine().sendLineScoreUpdateNotificationMenu(); }
function previewClassStatistics() { getEngine().previewClassStatistics(); }
function autoFormatAndResizeAllSheets() { getEngine().autoFormatAndResizeAllSheets(); }
function clearSystemCache() {
  _cachedEngine = null;
  getEngine(true).clearSystemCache();
}

// =========================================================================
// 試算表即時編輯原生監聽器（方案 B & C 本地執行，無需外部網路請求，確保 100% 穩定觸發）
// =========================================================================
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    if (isSystemSheet_(sheetName)) return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (col < 3) return;

    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var now = new Date();
    var nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    var nowDateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

    var itemTitle = '';
    try {
      itemTitle = (sheet.getRange(3, col).getValue() || sheet.getRange(2, col).getValue() || ('項目' + (col - 2))).toString().trim();
    } catch (tErr) {}

    // 情況 1：編輯項目名稱（第 3 列，例如新輸入 L6測驗卷）
    if (row === 3) {
      var newTitle = (e.value !== undefined ? e.value : sheet.getRange(row, col).getValue() || '').toString().trim();
      if (newTitle) {
        setupSchemeCColumnHeader_(sheet, col, newTitle);
        logScoreUpdateRecord_(sheetName, newTitle, '新增/設定評量項目', '第 ' + col + ' 欄', '試算表即時編輯');
      }
    }
    // 情況 2：編輯學生成績（第 4 列起）
    else if (row >= 4) {
      if (!itemTitle) itemTitle = '項目' + (col - 2);
      setupSchemeCColumnHeader_(sheet, col, itemTitle);
      try {
        var headerCell = sheet.getRange(3, col);
        var oldNote = headerCell.getNote() || '';
        var cleanNote = oldNote.replace(/最後更新：[^\n]+/g, '').trim();
        headerCell.setNote((cleanNote ? cleanNote + '\n' : '') + '最後更新：' + nowStr);
      } catch (notErr) {}

      try {
        var docProps = PropertiesService.getDocumentProperties();
        docProps.setProperty('EDIT_' + sheetName + '_' + col, nowStr);
      } catch (prpErr) {}

      try {
        var studentSeat = (sheet.getRange(row, 1).getValue() || '').toString().trim();
        var studentName = (sheet.getRange(row, 2).getValue() || '').toString().trim();
        var studentDesc = (studentSeat ? (studentSeat + '號 ') : '') + studentName;
        var editVal = e.value !== undefined ? e.value : sheet.getRange(row, col).getValue();
        var scoreDesc = studentDesc ? (studentDesc + '（' + editVal + '分）') : ('數值：' + editVal);
        logScoreUpdateRecord_(sheetName, itemTitle, '成績登錄/修改', scoreDesc, '試算表即時編輯');
      } catch (logErr) {}
    }
    // 情況 3：手動修改或從日曆選取器挑選第 1 列「登記日期」
    else if (row === 1) {
      var newDate = e.value !== undefined ? e.value : sheet.getRange(row, col).getValue();
      setupSchemeCColumnHeader_(sheet, col, itemTitle, newDate);
      logScoreUpdateRecord_(sheetName, itemTitle, '變更評量日期', '設為 ' + newDate, '試算表即時編輯');
    }
    // 情況 4：從下拉選單切換第 2 列「評量類別（作業/考試/不計分）」
    else if (row === 2) {
      var newType = e.value !== undefined ? e.value : sheet.getRange(row, col).getValue();
      setupSchemeCColumnHeader_(sheet, col, itemTitle, null, newType);
      logScoreUpdateRecord_(sheetName, itemTitle, '變更評量類別', '設為 ' + newType, '試算表即時編輯');
    }
  } catch (err) {}
}

/**
 * ➕ 新增 1 欄評量項目（方案 C 專用：自動在右側建立新欄、填入今日日期、日曆選取器與作業/考試下拉選單）
 */
function addNewAssessmentColumnMenu() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();
    if (isSystemSheet_(sheetName)) {
      SpreadsheetApp.getUi().alert('⚠️ 請先切換至欲新增評量的「科目分頁」（如國語、數學），再執行此功能。');
      return;
    }

    var lastCol = sheet.getLastColumn();
    var targetCol = Math.max(3, lastCol + 1);

    setupSchemeCColumnHeader_(sheet, targetCol, '新評量項目');
    sheet.setActiveRange(sheet.getRange(3, targetCol));
    SpreadsheetApp.getActiveSpreadsheet().toast('已成功於第 ' + targetCol + ' 欄建立新評量欄位！已自動填入今日日期與作業/考試下拉選單。', '🎉 新增評量欄位成功', 4);
  } catch (err) {
    SpreadsheetApp.getUi().alert('新增欄位失敗：' + err.toString());
  }
}

function isSystemSheet_(sheetName) {
  if (!sheetName) return true;
  if (sheetName === '學生名冊' || sheetName === '家長閱覽紀錄' || sheetName === '成績異動日誌' || sheetName.startsWith('_')) return true;
  if (sheetName.indexOf('總表') !== -1 || sheetName.indexOf('總成績') !== -1) return true;
  if (sheetName.indexOf('閱覽') !== -1 || sheetName.indexOf('紀錄') !== -1 || sheetName.indexOf('日誌') !== -1) return true;
  return false;
}

function inferDefaultAssessmentType_(itemName) {
  if (!itemName) return '作業';
  var h = itemName.toString().trim();
  if (h.indexOf('考') !== -1 || h.indexOf('測') !== -1 || h.indexOf('卷') !== -1 || h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1 || h.indexOf('生字') !== -1) {
    return '考試';
  }
  return '作業';
}

function setupSchemeCColumnHeader_(sheet, colIdx, optItemTitle, optDateVal, optTypeVal) {
  if (!sheet || colIdx < 3) return;
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var now = new Date();
    var nowDateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

    var titleCell = sheet.getRange(3, colIdx);
    var itemTitle = (optItemTitle !== undefined && optItemTitle !== null ? optItemTitle : titleCell.getValue() || '').toString().trim();
    if (!itemTitle) {
      var lr = sheet.getLastRow();
      if (lr >= 4) {
        var colVals = sheet.getRange(4, colIdx, Math.min(lr - 3, 35), 1).getValues();
        for (var vi = 0; vi < colVals.length; vi++) {
          if (colVals[vi][0] !== '' && colVals[vi][0] !== null && colVals[vi][0] !== undefined) {
            itemTitle = '項目' + (colIdx - 2);
            titleCell.setValue(itemTitle);
            break;
          }
        }
      }
    }

    var curDateRaw = sheet.getRange(1, colIdx).getValue();
    var curTypeRaw = sheet.getRange(2, colIdx).getValue();
    if (!itemTitle && !curDateRaw && !curTypeRaw && optDateVal === undefined && optTypeVal === undefined && optItemTitle === undefined) {
      return;
    }

    var dateCell = sheet.getRange(1, colIdx);
    var curDateVal = (optDateVal !== undefined && optDateVal !== null ? optDateVal : dateCell.getValue() || '').toString().trim();
    if (!curDateVal) {
      curDateVal = nowDateStr;
      dateCell.setValue(curDateVal);
    }
    dateCell.setNumberFormat('yyyy/MM/dd')
      .setBackground('#f0fdf4')
      .setFontColor('#166534')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    try {
      var dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
      dateCell.setDataValidation(dateRule);
    } catch (dValErr) {}

    var typeCell = sheet.getRange(2, colIdx);
    var curTypeVal = (optTypeVal !== undefined && optTypeVal !== null ? optTypeVal : typeCell.getValue() || '').toString().trim();
    if (!curTypeVal || (curTypeVal !== '作業' && curTypeVal !== '考試' && curTypeVal !== '不計分')) {
      curTypeVal = inferDefaultAssessmentType_(itemTitle);
      typeCell.setValue(curTypeVal);
    }

    try {
      var typeRule = SpreadsheetApp.newDataValidation().requireValueInList(['作業', '考試', '不計分'], true).setAllowInvalid(true).build();
      typeCell.setDataValidation(typeRule);
    } catch (tValErr) {}

    if (curTypeVal === '作業') {
      typeCell.setBackground('#fef3c7').setFontColor('#92400e');
    } else if (curTypeVal === '考試') {
      typeCell.setBackground('#e0f2fe').setFontColor('#075985');
    } else {
      typeCell.setBackground('#f5f5f4').setFontColor('#78716c');
    }
    typeCell.setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

    if (itemTitle) {
      if (itemTitle.indexOf('習') !== -1) {
        titleCell.setBackground('#d97706').setFontColor('#ffffff');
      } else if (curTypeVal === '作業' || itemTitle.indexOf('作業') !== -1 || itemTitle.indexOf('學習單') !== -1) {
        titleCell.setBackground('#059669').setFontColor('#ffffff');
      } else if (itemTitle.indexOf('語詞') !== -1 || itemTitle.indexOf('聽寫') !== -1 || itemTitle.indexOf('圈詞') !== -1 || itemTitle.indexOf('生字') !== -1) {
        titleCell.setBackground('#0284c7').setFontColor('#ffffff');
      } else {
        titleCell.setBackground('#b45309').setFontColor('#ffffff');
      }
      titleCell.setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    sheet.setColumnWidth(colIdx, 84);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 26);
    sheet.setRowHeight(3, 32);
  } catch (err) {}
}

function logScoreUpdateRecord_(subjectName, assessmentName, updateType, studentInfo, remark) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('成績異動日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('成績異動日誌');
      var headers = [['時間戳記', '科目分頁', '評量項目', '異動類型', '影響學生/數值', '操作來源/備註']];
      logSheet.getRange(1, 1, 1, 6).setValues(headers)
        .setBackground('#0f172a')
        .setFontColor('#f8fafc')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      logSheet.setColumnWidth(1, 160);
      logSheet.setColumnWidth(2, 120);
      logSheet.setColumnWidth(3, 140);
      logSheet.setColumnWidth(4, 110);
      logSheet.setColumnWidth(5, 140);
      logSheet.setColumnWidth(6, 200);
      try {
        logSheet.setFrozenRows(1);
      } catch (fErr) {}
    }

    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([
      nowStr,
      subjectName || '',
      assessmentName || '',
      updateType || '成績登記',
      studentInfo || '',
      remark || '試算表即時編輯'
    ]);
  } catch (err) {}
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

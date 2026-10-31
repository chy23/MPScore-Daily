/**
 * MPScore-Daily - 學生平時成績歷程與診斷系統 (Google Apps Script 後端核心)
 * 網站作者 / 系統規劃：楊家驊老師 (Chiahua Yang)
 *
 * 【核心架構 — 專為各科平時成績（習作、作業卷、語詞測驗、單元測驗卷）統計與歷程分析設計】：
 * 1. 【學生名冊】分頁：班級 | 座號 | 姓名 | 專屬代碼 | 啟用狀態 (可與 MPScore 段考系統 100% 共用代碼)
 * 2. 各科平時成績分頁（例如：國語平時、數學平時、社會平時、國語期中平時、國語期末平時...）：
 *    - 支援「單列標題」（如：國習L1、作業卷L1、語詞L1、測驗卷L1）
 *    - 亦無痛支援老師原有的「雙列合併群組標題」（上列：國習 / 作業卷 / 語詞測驗 / 測驗卷；下列：L1、L2、L3...）
 *    - 自動略過試算表內原有的「平均、總均、平時成績、名次」等公式欄，由系統精準重算各分項平均、總平均、PR常模與單元走勢！
 */

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
    .addToUi();
}

/**
 * 導向至成績異動日誌分頁（方案 B）
 */
function openScoreUpdateLogsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('成績異動日誌');
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  } else {
    SpreadsheetApp.getUi().alert('提示', '目前尚無成績異動紀錄。當有成績登記、上傳或修改時，系統將自動建立【成績異動日誌】分頁！', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 清空成績異動日誌（方案 B 選單呼叫）
 */
function clearScoreUpdateLogsMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('確認清空成績異動日誌', '確定要清空【成績異動日誌】的所有紀錄嗎？', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName('成績異動日誌');
    if (logSheet && logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).clearContent();
    }
    ui.alert('成功', '✅ 【成績異動日誌】紀錄已清空！', ui.ButtonSet.OK);
  }
}

/**
 * 寫入成績異動日誌（方案 B 核心紀錄）
 */
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
      logSheet.setColumnWidth(1, 160); // 時間戳記
      logSheet.setColumnWidth(2, 120); // 科目分頁
      logSheet.setColumnWidth(3, 140); // 評量項目
      logSheet.setColumnWidth(4, 110); // 異動類型
      logSheet.setColumnWidth(5, 140); // 影響學生/數值
      logSheet.setColumnWidth(6, 200); // 操作來源/備註
      safeFreezeSheet_(logSheet, 1, 0);
    }

    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([
      nowStr,
      subjectName || '',
      assessmentName || '',
      updateType || '成績登記',
      studentInfo || '',
      remark || '系統自動紀錄'
    ]);
  } catch (err) {
    Logger.log('寫入成績異動日誌失敗: ' + err.toString());
  }
}

/**
 * 導向至家長閱覽紀錄分頁
 */
function openAccessLogsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  if (logSheet) {
    ss.setActiveSheet(logSheet);
  } else {
    SpreadsheetApp.getUi().alert('提示', '目前尚無家長閱覽紀錄。當有家長掃碼查看時，系統將自動建立【家長閱覽紀錄】分頁！', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 清空家長閱覽紀錄與次數重置（選單呼叫）
 */
function clearAllAccessLogsMenu() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert('確認清空家長閱覽紀錄', '確定要將全班家長閱覽次數歸零，並清空【家長閱覽紀錄】分頁嗎？\n\n（此動作適合在測試完畢後、正式發送小卡給家長前執行，將學生名冊閱覽次數與時間全數歸零）', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    clearAccessLogs_(ss, 'all');
    clearSystemCache();
    ui.alert('成功', '✅ 全班家長閱覽紀錄已清空，學生名冊累積次數與最後時間已全數歸零！', ui.ButtonSet.OK);
  }
}

/**
 * 清空家長閱覽紀錄與次數重置核心函式
 * @param {Spreadsheet} ss 
 * @param {string} scope 'all' | 'single'
 * @param {string} targetToken (optional, used when scope === 'single')
 */
function clearAccessLogs_(ss, scope, targetToken) {
  var sSheet = ss.getSheetByName('學生名冊');
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  var cache = CacheService.getScriptCache();
  var tokToClear = (targetToken || '').toString().trim().toUpperCase();

  // 1. 清空/重置【學生名冊】之 F 欄（最後閱覽時間）與 G 欄（累積閱覽次數）
  if (sSheet && sSheet.getLastRow() > 1) {
    var sData = sSheet.getDataRange().getValues();
    var lastCol = sSheet.getLastColumn();
    for (var r = 1; r < sData.length; r++) {
      var rowTok = (sData[r][3] || '').toString().trim().toUpperCase();
      if (scope === 'all' || (scope === 'single' && rowTok === tokToClear)) {
        if (lastCol >= 6) sSheet.getRange(r + 1, 6).setValue('');
        if (lastCol >= 7) sSheet.getRange(r + 1, 7).setValue(0);
        if (rowTok) {
          try {
            cache.remove('mpdaily_visit_debounce_' + rowTok);
          } catch (ignore) {}
        }
      }
    }
  }

  // 2. 清空【家長閱覽紀錄】專屬日誌
  if (logSheet && logSheet.getLastRow() > 1) {
    if (scope === 'all' || !tokToClear) {
      var numRows = logSheet.getLastRow() - 1;
      logSheet.deleteRows(2, numRows);
    } else {
      var logData = logSheet.getDataRange().getValues();
      for (var i = logData.length - 1; i >= 1; i--) {
        var rowTok = (logData[i][4] || '').toString().trim().toUpperCase();
        if (rowTok === tokToClear) {
          logSheet.deleteRow(i + 1);
        }
      }
    }
  }

  return { success: true };
}

/**
 * 取得當前教師安全金鑰 (PIN)
 */
function getTeacherAdminKey_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('TEACHER_ADMIN_KEY');
  if (key && key.trim()) return key.trim();

  // 若屬性尚未設定，檢查【學生名冊】之 J1 儲存格
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    if (sSheet && sSheet.getLastRow() >= 1) {
      var val = (sSheet.getRange(1, 10).getValue() || '').toString().trim();
      if (val && val !== '【教師安全金鑰】') {
        props.setProperty('TEACHER_ADMIN_KEY', val);
        return val;
      }
    }
  } catch (e) {}

  // 預設金鑰（初次使用無痛過渡）
  return '8888';
}

/**
 * 寫入/更新教師安全金鑰 (PIN)
 */
function setTeacherAdminKey_(newKey) {
  if (!newKey || !newKey.toString().trim()) return false;
  var k = newKey.toString().trim();
  PropertiesService.getScriptProperties().setProperty('TEACHER_ADMIN_KEY', k);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    if (sSheet) {
      sSheet.getRange(1, 9).setValue('【教師安全金鑰】').setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#334155').setHorizontalAlignment('center');
      sSheet.getRange(1, 10).setValue(k).setFontWeight('bold').setBackground('#e0f2fe').setFontColor('#0369a1').setHorizontalAlignment('center');
    }
  } catch (e) {}
  return true;
}

/**
 * 驗證傳入之請求是否具備合法教師安全金鑰
 */
function verifyTeacherAuth_(e) {
  var params = e ? e.parameter : {};
  var clientKey = (params.key || params.adminKey || params.teacherKey || params.pin || '').toString().trim();
  var serverKey = getTeacherAdminKey_();

  // 支援教師在 J1 儲存格直接修改試算表
  var sheetKey = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    if (sSheet && sSheet.getLastRow() >= 1) {
      sheetKey = (sSheet.getRange(1, 10).getValue() || '').toString().trim();
    }
  } catch (e) {}

  if (!clientKey) return false;
  if (clientKey === serverKey) return true;
  if (sheetKey && clientKey === sheetKey) {
    PropertiesService.getScriptProperties().setProperty('TEACHER_ADMIN_KEY', sheetKey);
    return true;
  }
  return false;
}

/**
 * 🛡️ 設定/查看「教師安全金鑰 (PIN)」選單對話框
 */
function manageTeacherAdminKeyMenu() {
  var ui = SpreadsheetApp.getUi();
  var currentKey = getTeacherAdminKey_();
  var isDefault = (currentKey === '8888');

  var msg = '【目前教師安全金鑰 (PIN)】: ' + currentKey + (isDefault ? '（⚠️ 目前為預設值 8888，建議自訂）' : '（✅ 已啟用自訂防護）') +
            '\n\n💡 說明：\n此金鑰用於保護全班成績與名冊，防止未經授權者透過 API 批次讀取。網頁端列印小卡或匯入成績時需輸入此 PIN 碼。\n\n如欲修改金鑰，請在下方輸入 4～10 位自訂英數密碼（若留空並按確定則維持原金鑰）：';

  var res = ui.prompt('🛡️ 管理教師安全金鑰 (PIN)', msg, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    var input = res.getResponseText().trim();
    if (input) {
      if (input.length < 4) {
        ui.alert('⚠️ 金鑰長度至少需 4 碼！未做變更。');
        return;
      }
      setTeacherAdminKey_(input);
      clearSystemCache();
      ui.alert('✅ 設定成功！\n\n新教師安全金鑰已設為：' + input + '\n您可在【學生名冊】第 1 列 J1 查看，亦請在網頁教師端輸入此 PIN 碼以同步資料。');
    }
  }
}

/**
 * 安全設定凍結列與凍結欄（自動拆解跨越凍結線的合併儲存格，徹底杜絕「無法凍結僅包含部分合併儲存格的欄」例外錯誤）
 */
function safeFreezeSheet_(sheet, freezeRows, freezeCols) {
  try {
    if (freezeRows > 0) {
      sheet.setFrozenRows(freezeRows);
    }
  } catch (e) {}

  if (freezeCols > 0) {
    try {
      sheet.setFrozenColumns(freezeCols);
    } catch (err) {
      // 若前幾列有跨欄合併儲存格（例如第 1 列大標題 A1:Z1 合併），先將跨越 freezeCols 的合併拆開為左、右兩塊再凍結
      try {
        var maxCol = Math.max(freezeCols + 1, sheet.getLastColumn());
        var checkRows = Math.max(2, freezeRows || 2);
        var topRange = sheet.getRange(1, 1, checkRows, maxCol);
        var mergedRanges = topRange.getMergedRanges();
        for (var i = 0; i < mergedRanges.length; i++) {
          var mr = mergedRanges[i];
          var startCol = mr.getColumn();
          var endCol = mr.getLastColumn();
          var startRow = mr.getRow();
          var numRows = mr.getNumRows();
          if (startCol <= freezeCols && endCol > freezeCols) {
            var originalVal = mr.getCell(1, 1).getValue();
            var bg = mr.getCell(1, 1).getBackground();
            var fc = mr.getCell(1, 1).getFontColor();
            var fw = mr.getCell(1, 1).getFontWeight();
            var fs = mr.getCell(1, 1).getFontSize();
            mr.breakApart();

            // 左側凍結區（A1:B1）
            var leftRange = sheet.getRange(startRow, startCol, numRows, freezeCols - startCol + 1);
            if (leftRange.getNumColumns() > 1 || numRows > 1) leftRange.merge();
            leftRange.setValue('📌 學生名單')
              .setBackground(bg)
              .setFontColor(fc)
              .setFontWeight(fw)
              .setFontSize(fs)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');

            // 右側成績區（C1:最後一欄）
            var rightCols = endCol - freezeCols;
            if (rightCols >= 1) {
              var rightRange = sheet.getRange(startRow, freezeCols + 1, numRows, rightCols);
              if (rightCols > 1 || numRows > 1) rightRange.merge();
              rightRange.setValue(originalVal)
                .setBackground(bg)
                .setFontColor(fc)
                .setFontWeight(fw)
                .setFontSize(fs)
                .setHorizontalAlignment('center')
                .setVerticalAlignment('middle');
            }
          }
        }
        sheet.setFrozenColumns(freezeCols);
      } catch (ignore) {}
    }
  }
}

/**
 * 清除系統雲端快取
 */
function clearSystemCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cache = CacheService.getScriptCache();
    var keys = [
      'mpdaily_print_summary_v2',
      'mpdaily_print_summary_v3',
      'mpdaily_print_summary_v4',
      'mpdaily_print_summary_v5',
      'mpdaily_print_summary_v6',
      'mpdaily_print_summary_v7',
      'mpdaily_print_summary_v8'
    ];
    for (var i = 1; i <= 60; i++) {
      var s = i < 10 ? '0' + i : '' + i;
      ['v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8'].forEach(function(ver) {
        keys.push('mpdaily_score_' + ver + '_DEMO' + s);
        keys.push('mpdaily_score_' + ver + '_' + s);
      });
    }
    try {
      var sheets = ss.getSheets();
      for (var shIdx = 0; shIdx < sheets.length; shIdx++) {
        if (!isSystemSheet(sheets[shIdx].getName())) {
          ensureAssessmentTypeRow_(sheets[shIdx]);
        }
      }
      var sSheet = ss.getSheetByName('學生名冊');
      if (sSheet && sSheet.getLastRow() > 1) {
        var sData = sSheet.getDataRange().getValues();
        for (var r = 1; r < sData.length; r++) {
          var tok = (sData[r][3] || '').toString().trim().toUpperCase();
          if (tok) {
            ['v2', 'v3', 'v4', 'v5', 'v6', 'v7', 'v8'].forEach(function(ver) {
              keys.push('mpdaily_score_' + ver + '_' + tok);
            });
          }
        }
      }
    } catch (ignore) {}
    cache.removeAll(keys);

    // 確保【學生名冊】第 1 列 I1:J1 具備【教師安全金鑰】與 LINE 連線設定
    var currentKey = getTeacherAdminKey_();
    var sSheet = ss.getSheetByName('學生名冊');
    if (sSheet && sSheet.getLastRow() >= 1) {
      if ((sSheet.getRange(1, 9).getValue() || '').toString().trim() !== '【教師安全金鑰】') {
        sSheet.getRange(1, 9).setValue('【教師安全金鑰】').setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#334155').setHorizontalAlignment('center');
        sSheet.getRange(1, 10).setValue(currentKey).setFontWeight('bold').setBackground('#e0f2fe').setFontColor('#0369a1').setHorizontalAlignment('center');
        sSheet.setColumnWidth(8, 30);
        sSheet.setColumnWidth(9, 130);
        sSheet.setColumnWidth(10, 100);
      }
      initLineSettingsCellsOnRoster_(sSheet);
    }

    ui.alert('✨ 雲端快取清除成功！', '家長端與查詢系統現已載入試算表最新平時成績。\n\n🛡️【教師安全金鑰防護已就緒】\n目前金鑰為：' + currentKey + '\n已登記於【學生名冊】第 1 列 J1（您可隨時直接在 J1 儲存格修改密碼）。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('提示', '快取已同步重置。' + err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 座號標準化 (1 -> '01')
 */
function normalizeSeatNumber(val) {
  if (val === null || val === undefined) return '';
  var str = val.toString().trim();
  if (!str) return '';
  var num = parseInt(str, 10);
  if (!isNaN(num) && num > 0 && num < 100 && /^\d+$/.test(str)) {
    return num < 10 ? '0' + num : '' + num;
  }
  return str;
}

/**
 * 🛠️ 建立初始表格（空白試算表第一次使用）
 * 若【學生名冊】已有老師輸入的真實學生名單，會直接採用【學生名冊】的學生來建立各科平時成績表，絕不覆蓋真實名單！
 */
function setupStandardSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. 建立或讀取【學生名冊】
  var sSheet = ss.getSheetByName('學生名冊') || ss.getSheets()[0];
  sSheet.setName('學生名冊');
  var sHeaders = [['班級', '座號', '姓名', '專屬代碼', '啟用狀態', '最後閱覽時間', '累積閱覽次數']];
  sSheet.getRange(1, 1, 1, 7).setValues(sHeaders)
    .setBackground('#b45309')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sSheet.setColumnWidth(6, 150);
  sSheet.setColumnWidth(7, 110);

  // 初始化或顯示【教師安全金鑰】於第 1 列 I1:J1，以及 LINE 設定於 L1:O1
  var currentKey = getTeacherAdminKey_();
  sSheet.getRange(1, 9).setValue('【教師安全金鑰】').setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#334155').setHorizontalAlignment('center');
  sSheet.getRange(1, 10).setValue(currentKey).setFontWeight('bold').setBackground('#e0f2fe').setFontColor('#0369a1').setHorizontalAlignment('center');
  sSheet.setColumnWidth(8, 30);
  sSheet.setColumnWidth(9, 130);
  sSheet.setColumnWidth(10, 100);
  initLineSettingsCellsOnRoster_(sSheet);

  var defaultSampleRoster = [
    ['501', '01', '王小華', 'DEMO01', true],
    ['501', '02', '陳品妍', 'DEMO02', true],
    ['501', '03', '林宥廷', 'DEMO03', true],
    ['501', '04', '張語恩', 'DEMO04', true],
    ['501', '05', '李冠宇', 'DEMO05', true],
    ['501', '06', '黃子軒', 'DEMO06', true],
    ['501', '07', '吳采潔', 'DEMO07', true],
    ['501', '08', '蔡承翰', 'DEMO08', true],
    ['501', '09', '許芷晴', 'DEMO09', true],
    ['501', '10', '鄭宇翔', 'DEMO10', true]
  ];

  var activeRoster = [];
  if (sSheet.getLastRow() <= 1) {
    sSheet.getRange(2, 1, defaultSampleRoster.length, 5).setValues(defaultSampleRoster);
    activeRoster = defaultSampleRoster;
  } else {
    // 讀取【學生名冊】目前已有的真實學生名單，避免用範例名單蓋掉老師的學生！
    var existingData = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 5).getValues();
    for (var r = 0; r < existingData.length; r++) {
      var cls = (existingData[r][0] || '501').toString().trim();
      var seat = normalizeSeatNumber(existingData[r][1] || (r + 1));
      var name = (existingData[r][2] || '').toString().trim();
      var tok = (existingData[r][3] || '').toString().trim();
      var en = existingData[r][4] !== false;
      if (name) {
        activeRoster.push([cls, seat, name, tok, en]);
      }
    }
    if (activeRoster.length === 0) {
      sSheet.getRange(2, 1, defaultSampleRoster.length, 5).setValues(defaultSampleRoster);
      activeRoster = defaultSampleRoster;
    }
  }

  sSheet.setColumnWidth(1, 90);
  sSheet.setColumnWidth(2, 80);
  sSheet.setColumnWidth(3, 120);
  sSheet.setColumnWidth(4, 130);
  sSheet.setColumnWidth(5, 100);
  sSheet.setRowHeight(1, 36);
  var sLastRow = Math.max(2, sSheet.getLastRow());
  for (var sr = 2; sr <= sLastRow; sr++) {
    sSheet.setRowHeight(sr, 28);
  }
  sSheet.getRange(1, 1, sLastRow, 5)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  safeFreezeSheet_(sSheet, 1, 0);

  // 2. 建立【國語平時】分頁（若已存在則只更新名單不覆蓋既有成績）
  var chiHeaders = [
    '座號', '姓名',
    '國習L1', '國習L2', '國習L3', '國習L4', '國習L5', '國習L6',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5', '作業卷L6',
    '語詞L1', '語詞L2', '語詞L3', '語詞L4', '語詞L5', '語詞L6',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5', '測驗卷L6'
  ];
  var chiRows = activeRoster.map(function(st, idx) {
    var base = idx % 5 === 0 ? 94 : (idx % 5 === 1 ? 88 : (idx % 5 === 2 ? 91 : (idx % 5 === 3 ? 78 : 84)));
    return [
      st[1], st[2],
      Math.min(100, base + 4), Math.min(100, base + 5), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 6), Math.min(100, base + 5),
      Math.min(100, base + 3), Math.min(100, base + 4), Math.max(68, base - 4), Math.min(100, base + 5), Math.min(100, base + 3), Math.min(100, base + 4),
      Math.min(100, base + 2), Math.min(100, base), Math.max(62, base - 8), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 2),
      Math.min(100, base + 2), Math.min(100, base + 4), Math.max(58, base - 14), Math.max(66, base - 6), Math.min(100, base + 2), Math.min(100, base + 4)
    ];
  });
  var chiSheet = ss.getSheetByName('國語平時');
  if (!chiSheet) {
    chiSheet = ss.insertSheet('國語平時');
    formatDailySubjectSheet(chiSheet, '115學年度 國語科平時成績紀錄', chiHeaders, chiRows);
  }

  // 3. 建立【數學平時】分頁
  var mathHeaders = [
    '座號', '姓名',
    '數習L1', '數習L2', '數習L3', '數習L4', '數習L5', '數習L6',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5', '作業卷L6',
    '計算L1', '計算L2', '計算L3', '計算L4', '計算L5', '計算L6',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5', '測驗卷L6'
  ];
  var mathRows = activeRoster.map(function(st, idx) {
    var base = idx % 4 === 0 ? 90 : (idx % 4 === 1 ? 85 : (idx % 4 === 2 ? 93 : 76));
    return [
      st[1], st[2],
      Math.min(100, base + 5), Math.min(100, base + 4), Math.min(100, base + 2), Math.min(100, base + 6), Math.min(100, base + 5), Math.min(100, base + 4),
      Math.min(100, base + 2), Math.min(100, base + 3), Math.max(64, base - 6), Math.min(100, base + 4), Math.min(100, base + 3), Math.min(100, base + 4),
      Math.min(100, base + 4), Math.min(100, base + 5), Math.max(70, base - 4), Math.min(100, base + 4), Math.min(100, base + 6), Math.min(100, base + 5),
      Math.min(100, base + 3), Math.max(60, base - 12), Math.min(100, base), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 3)
    ];
  });
  var mathSheet = ss.getSheetByName('數學平時');
  if (!mathSheet) {
    mathSheet = ss.insertSheet('數學平時');
    formatDailySubjectSheet(mathSheet, '115學年度 數學科平時成績紀錄', mathHeaders, mathRows);
  }

  // 4. 建立【社會平時】分頁
  var socHeaders = [
    '座號', '姓名',
    '社習L1', '社習L2', '社習L3', '社習L4', '社習L5', '社習L6',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5', '作業卷L6',
    '重點L1', '重點L2', '重點L3', '重點L4', '重點L5', '重點L6',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5', '測驗卷L6'
  ];
  var socRows = activeRoster.map(function(st, idx) {
    var base = idx % 4 === 0 ? 92 : (idx % 4 === 1 ? 88 : (idx % 4 === 2 ? 90 : 80));
    return [
      st[1], st[2],
      Math.min(100, base + 4), Math.min(100, base + 6), Math.min(100, base + 3), Math.min(100, base + 5), Math.min(100, base + 4), Math.min(100, base + 6),
      Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 1), Math.min(100, base + 3), Math.min(100, base + 2), Math.min(100, base + 5),
      Math.min(100, base + 3), Math.min(100, base + 5), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 3), Math.min(100, base + 5),
      Math.min(100, base), Math.min(100, base + 2), Math.max(68, base - 8), Math.min(100, base + 4), Math.min(100, base + 2), Math.min(100, base + 5)
    ];
  });
  var socSheet = ss.getSheetByName('社會平時');
  if (!socSheet) {
    socSheet = ss.insertSheet('社會平時');
    formatDailySubjectSheet(socSheet, '115學年度 社會科平時成績紀錄', socHeaders, socRows);
  }

  // 確保【學生名冊】的姓名也同步到各科分頁
  syncRosterToSubjectSheets(true);

  ui.alert(
    '🎉 平時成績表格建置完成！\n\n' +
    '已為您備妥：\n' +
    '1. 【學生名冊】(共 ' + activeRoster.length + ' 位學生，可與段考系統共用代碼)\n' +
    '2. 【國語平時】、【數學平時】、【社會平時】\n\n' +
    '💡 貼心提醒：\n' +
    '・若您修改了【學生名冊】的學生姓名，請點選選單「📤 將『學生名冊』姓名同步至各科平時成績表」。\n' +
    '・若您是將新名單貼在某個科目分頁，請切換到該分頁後點選「📥 從目前成績分頁匯入名單至『學生名冊』」。'
  );
}

/**
 * 格式化平時成績科目分頁（方案 C 專用規格）
 * - 第 1 列：評量登記/更新日期（A1:B1 為「📌 登記日期」，C1.. 為各項目的日期 YYYY/MM/DD）
 * - 第 2 列：考試或作業的選項（評量類別：作業 / 考試 / 不計分 下拉選單）
 * - 第 3 列：項目名稱標題列（座號、姓名、國習L1、作業卷L1...）
 * - 第 4 列起：學生成績資料列
 */
function formatDailySubjectSheet(sheet, titleText, headers, rowsData) {
  sheet.clear();
  var numCols = Math.max(3, headers.length);
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var todayDateStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

  // --- 第 1 列：日期列 ---
  // A1:B1 合併顯示「📌 登記日期」
  sheet.getRange(1, 1, 1, 2).merge()
    .setValue('📌 登記日期')
    .setBackground('#1e293b')
    .setFontColor('#e2e8f0')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // C1 到最後一欄：設定各評量欄位的初始日期（預設為今天 YYYY/MM/DD）
  var dateValues = [];
  var dateBgs = [];
  var dateFontColors = [];
  for (var dc = 3; dc <= numCols; dc++) {
    var headerTitle = (headers[dc - 1] || '').toString().trim();
    var parsedD = extractDateFromTitle_(headerTitle);
    var dStr = parsedD ? Utilities.formatDate(parsedD, tz, 'yyyy/MM/dd') : todayDateStr;
    dateValues.push(dStr);
    dateBgs.push('#f0fdf4');
    dateFontColors.push('#166534');
  }
  if (dateValues.length > 0) {
    var dateRange = sheet.getRange(1, 3, 1, dateValues.length);
    dateRange
      .setValues([dateValues])
      .setBackgrounds([dateBgs])
      .setFontColors([dateFontColors])
      .setFontWeight('bold')
      .setFontSize(10)
      .setNumberFormat('yyyy/MM/dd')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 啟用 Google 試算表原生雙擊/點選彈出日曆選取器
    try {
      var dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .build();
      dateRange.setDataValidation(dateRule);
    } catch (dValErr) {}
  }
  sheet.setRowHeight(1, 26);

  // --- 第 2 列：作業 / 考試 選項列 ---
  var typeRowValues = ['評量', '作業/考試'];
  var typeBgColors = ['#475569', '#475569'];
  var typeFontColors = ['#ffffff', '#ffffff'];

  for (var tc = 3; tc <= numCols; tc++) {
    var rawH = (headers[tc - 1] || '').toString().trim();
    var defaultType = inferDefaultAssessmentType(rawH);
    typeRowValues.push(defaultType);
    if (defaultType === '作業') {
      typeBgColors.push('#fef3c7');
      typeFontColors.push('#92400e');
    } else if (defaultType === '考試') {
      typeBgColors.push('#e0f2fe');
      typeFontColors.push('#075985');
    } else {
      typeBgColors.push('#f5f5f4');
      typeFontColors.push('#78716c');
    }
  }

  sheet.getRange(2, 1, 1, numCols)
    .setValues([typeRowValues])
    .setBackgrounds([typeBgColors])
    .setFontColors([typeFontColors])
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(2, 26);

  if (numCols > 2) {
    try {
      var validationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['作業', '考試', '不計分'], true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 3, 1, numCols - 2).setDataValidation(validationRule);
    } catch (ignore) {}
  }

  // --- 第 3 列：項目名稱標題列 ---
  var itemHeaders = headers.slice(0, numCols);
  sheet.getRange(3, 1, 1, numCols).setValues([itemHeaders])
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(3, 32);

  for (var c = 1; c <= numCols; c++) {
    var h = (headers[c - 1] || '').toString().trim();
    var itemCell = sheet.getRange(3, c);
    if (c === 1 || c === 2) {
      itemCell.setBackground('#334155').setFontColor('#ffffff');
    } else {
      var t = typeRowValues[c - 1];
      if (h.indexOf('習') !== -1) {
        itemCell.setBackground('#d97706').setFontColor('#ffffff');
      } else if (t === '作業') {
        itemCell.setBackground('#059669').setFontColor('#ffffff');
      } else if (h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1) {
        itemCell.setBackground('#0284c7').setFontColor('#ffffff');
      } else {
        itemCell.setBackground('#b45309').setFontColor('#ffffff');
      }
    }
  }

  // --- 第 4 列起：學生成績資料 ---
  if (rowsData && rowsData.length > 0) {
    sheet.getRange(4, 1, rowsData.length, numCols).setValues(rowsData)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  sheet.setColumnWidth(1, 70);  // 座號
  sheet.setColumnWidth(2, 110); // 姓名
  for (var col = 3; col <= numCols; col++) {
    sheet.setColumnWidth(col, 84);
  }
  var lastRow = Math.max(3, sheet.getLastRow());
  for (var r = 4; r <= lastRow; r++) {
    sheet.setRowHeight(r, 28);
  }
  sheet.getRange(1, 1, lastRow, numCols).setVerticalAlignment('middle').setHorizontalAlignment('center');

  safeFreezeSheet_(sheet, 3, 2);

  // 方案 B：寫入成績異動日誌
  logScoreUpdateRecord_(sheet.getName(), '全部分頁評量', '格式化建立分頁', (rowsData ? rowsData.length : 0) + ' 位學生', '方案 C 版面');
}

/**
 * Web API 入口 (doGet / doPost)
 */
function doGet(e) {
  var params = e ? e.parameter : {};
  var action = params.action || (params.token ? 'query' : 'ping');
  var output = {};

  try {
    if (action === 'ping') {
      output = checkSystemStatus();
    } else if (action === 'query') {
      var isSimulated = (params.sim === '1' || params.preview === '1' || params.simulate === '1' || (params.device && params.device.indexOf('模擬') !== -1));
      output = getStudentDailyPortfolioByToken(params.token, params.device, (params.nocache === '1' || params.reload === '1'), isSimulated);
    } else if (action === 'getStudentLogs') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var studentToken = (params.token || '').toString().trim().toUpperCase();
      var historyMap = getStudentAccessHistoryMap_(ss);
      output = {
        status: 'success',
        success: true,
        token: studentToken,
        recentViews: historyMap[studentToken] || []
      };
    } else {
      // 🔒 安全升級：所有管理端資料操作皆強制實施「教師安全金鑰 (PIN)」驗證
      if (!verifyTeacherAuth_(e)) {
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          success: false,
          code: 'AUTH_FAILED',
          message: '🔒 存取遭拒：教師管理安全金鑰 (PIN) 錯誤或未提供！此操作僅限授權教師執行。'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      if (action === 'verifyKey' || action === 'checkAuth') {
        output = { status: 'success', success: true, message: '教師金鑰驗證通過' };
      } else if (action === 'getStudentsForPrint' || action === 'getRoster') {
        output = getStudentListForTeacher();
      } else if (action === 'getAccessLogs') {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        output = {
          status: 'success',
          success: true,
          updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm'),
          accessLogs: getRecentAccessLogs_(ss, 100)
        };
      } else if (action === 'clearAccessLogs') {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var scope = (params.scope || 'all').toString().trim();
        var studentToken = (params.token || '').toString().trim().toUpperCase();
        clearAccessLogs_(ss, scope, studentToken);
        output = {
          status: 'success',
          success: true,
          message: scope === 'all' ? '全班閱覽紀錄已清空歸零' : '該生閱覽紀錄已清除',
          scope: scope,
          token: studentToken
        };
      } else if (action === 'appendDailyScores') {
        var body = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
        output = appendDailyScores(body);
      } else if (action === 'getLineStatus' || action === 'getSubjectUpdateCounts') {
        output = getLinePushStatus_(params);
      } else if (action === 'sendLineNotification') {
        output = executeLineScorePushNotifications_(params);
      } else if (action === 'setTeacherKey') {
        var newKey = (params.newKey || '').toString().trim();
        if (newKey && newKey.length >= 4) {
          setTeacherAdminKey_(newKey);
          output = { status: 'success', success: true, message: '教師安全金鑰已成功更新！' };
        } else {
          output = { status: 'error', success: false, message: '新金鑰長度至少需 4 碼！' };
        }
      } else {
        output = { status: 'error', success: false, message: '未知的操作請求: ' + action };
      }
    }
  } catch (err) {
    output = { status: 'error', success: false, message: '系統處理錯誤: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

/**
 * 判斷是否為系統名冊分頁
 */
function isSystemSheet(sheetName) {
  if (!sheetName) return true;
  if (sheetName === '學生名冊' || sheetName === '家長閱覽紀錄' || sheetName.startsWith('_')) return true;
  if (sheetName.indexOf('總表') !== -1 || sheetName.indexOf('總成績') !== -1) return true;
  if (sheetName.indexOf('閱覽') !== -1 || sheetName.indexOf('紀錄') !== -1 || sheetName.indexOf('日誌') !== -1) return true;
  return false;
}

/**
 * 判斷欄位標題是否為試算表內既有的「統計結果欄或加分欄」（例如：平均、總均、平時成績、名次、資源班、加分），而非單次平時評量 0~100 分數
 */
function isSummaryColumnHeader(headerText) {
  if (!headerText) return false;
  var h = headerText.toString().trim();
  if (!h) return false;
  if (h === '座號' || h === '姓名' || h === '班級' || h === '學號' || h === '備註') return true;
  var summaryKeywords = ['平均', '總均', '總分', '平時成績', '平時總成績', '學期成績', '期中平時', '期末平時', '排名', '名次', '資源班', '加權', '百分比', '占比', '加分', '扣分'];
  for (var i = 0; i < summaryKeywords.length; i++) {
    if (h.indexOf(summaryKeywords[i]) !== -1) return true;
  }
  return false;
}

/**
 * 電腦初步判斷單項評量屬於「作業」或「考試」
 */
function inferDefaultAssessmentType(itemName) {
  var h = (itemName || '').toString().trim();
  if (!h) return '';
  if (isSummaryColumnHeader(h)) return '不計分';
  if (h.indexOf('習') !== -1 || h.indexOf('作業') !== -1 || h.indexOf('練習') !== -1 || h.indexOf('學習單') !== -1 || h.indexOf('訂正') !== -1 || h.indexOf('筆記') !== -1) {
    return '作業';
  }
  if (h.indexOf('測驗') !== -1 || h.indexOf('考') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1 || h.indexOf('卷') !== -1) {
    return '考試';
  }
  return '作業';
}

/**
 * 判斷某列是否為「作業／考試」類別設定列
 */
function isAssessmentTypeRow_(row, seatCol, nameCol) {
  if (!row || row.length === 0) return false;
  var sCell = (row[seatCol] !== undefined && row[seatCol] !== null) ? row[seatCol].toString().trim() : '';
  var nCell = (row[nameCol] !== undefined && row[nameCol] !== null) ? row[nameCol].toString().trim() : '';
  if (sCell === '評量' || sCell === '類別' || sCell === '類型' || sCell === '性質' || nCell === '作業/考試' || nCell === '類別' || nCell === '評量類別') {
    return true;
  }
  var isNotStudentSeat = (!sCell || isNaN(parseInt(sCell, 10)));
  if (isNotStudentSeat && !nCell) {
    var hasTypeKeyword = false;
    var hasNonTypeItem = false;
    for (var i = 0; i < row.length; i++) {
      if (i === seatCol || i === nameCol) continue;
      var v = (row[i] || '').toString().trim();
      if (!v) continue;
      if (v === '作業' || v === '考試' || v === '不計分' || v === '略過' || v === '習作' || v === '測驗' || v === '加分' || v === '扣分') {
        hasTypeKeyword = true;
      } else {
        hasNonTypeItem = true;
      }
    }
    if (hasTypeKeyword && !hasNonTypeItem) return true;
  }
  return false;
}

/**
 * 智慧定位科目平時成績表的各關鍵列位置（完整相容單列標題、雙列群組與垂直合併「座號/姓名」標題，並自動識別異常空白/錯位列）
 */
function locateSubjectSheetRows_(data) {
  var result = {
    seatCol: 0,
    nameCol: 1,
    dateRowIdx: -1,
    typeRowIdx: -1,
    itemHeaderRowIdx: -1,
    firstStudentRowIdx: -1,
    misplacedRowIndices: []
  };
  if (!data || data.length === 0) return result;

  var headerCandidateRowIdx = -1;
  for (var r = 0; r < Math.min(10, data.length); r++) {
    for (var c = 0; c < Math.min(8, data[r].length); c++) {
      var cellStr = (data[r][c] || '').toString().trim();
      if (cellStr === '座號' || cellStr.indexOf('座號') !== -1) {
        result.seatCol = c;
        if (headerCandidateRowIdx === -1) headerCandidateRowIdx = r;
      } else if (cellStr === '姓名' || cellStr.indexOf('姓名') !== -1) {
        result.nameCol = c;
        if (headerCandidateRowIdx === -1) headerCandidateRowIdx = r;
      }
    }
  }

  // 1. 尋找第一位學生的資料列 (具備有效座號數字 1..99 與學生姓名)
  for (var sRow = 1; sRow < data.length; sRow++) {
    if (isAssessmentTypeRow_(data[sRow], result.seatCol, result.nameCol)) continue;
    var sVal = (data[sRow][result.seatCol] !== undefined && data[sRow][result.seatCol] !== null) ? data[sRow][result.seatCol].toString().trim() : '';
    var nVal = (data[sRow][result.nameCol] !== undefined && data[sRow][result.nameCol] !== null) ? data[sRow][result.nameCol].toString().trim() : '';
    if (!sVal || !nVal) continue;
    if (/^\d{1,2}$/.test(sVal)) {
      var seatNum = parseInt(sVal, 10);
      if (seatNum >= 1 && seatNum <= 99 && nVal.indexOf('姓名') === -1 && nVal.indexOf('平均') === -1 && nVal.indexOf('作業') === -1 && nVal.indexOf('標準差') === -1) {
        result.firstStudentRowIdx = sRow;
        break;
      }
    }
  }

  var searchUpper = result.firstStudentRowIdx !== -1 ? result.firstStudentRowIdx : Math.min(data.length, 12);

  // 2. 尋找項目名稱標題列 (包含座號、姓名及評量項目如數習L1、測驗卷)
  if (headerCandidateRowIdx !== -1) {
    result.itemHeaderRowIdx = headerCandidateRowIdx;
  } else {
    var bestItemCount = -1;
    for (var hRow = 0; hRow < searchUpper; hRow++) {
      if (isAssessmentTypeRow_(data[hRow], result.seatCol, result.nameCol)) continue;
      var itemCount = 0;
      for (var col = 0; col < data[hRow].length; col++) {
        if (col === result.seatCol || col === result.nameCol) continue;
        var v = (data[hRow][col] || '').toString().trim();
        if (!v || v === '📌 學生名單' || v === '📌 登記日期') continue;
        if (v === '作業' || v === '考試' || v === '不計分' || v === '略過') continue;
        itemCount++;
      }
      if (itemCount > 0 && itemCount >= bestItemCount) {
        bestItemCount = itemCount;
        result.itemHeaderRowIdx = hRow;
      }
    }
  }

  if (result.itemHeaderRowIdx === -1) return result;

  // 3. 尋找「作業／考試」類別列 (支援方案 C 位於第 2 列，或傳統結構位於第 3 列)
  for (var tRow = 0; tRow < searchUpper; tRow++) {
    if (tRow === result.itemHeaderRowIdx) continue;
    if (isAssessmentTypeRow_(data[tRow], result.seatCol, result.nameCol)) {
      if (result.typeRowIdx === -1) {
        result.typeRowIdx = tRow;
      }
    }
  }

  // 4. 尋找「登記日期列」（方案 C 中位於第 1 列，通常含有「日期」字樣或日期格式）
  for (var dRow = 0; dRow < searchUpper; dRow++) {
    if (dRow === result.itemHeaderRowIdx || dRow === result.typeRowIdx) continue;
    var dCell = (data[dRow][result.seatCol] || data[dRow][0] || '').toString().trim();
    if (dCell.indexOf('日期') !== -1 || dRow === 0) {
      result.dateRowIdx = dRow;
      break;
    }
  }

  // 5. 檢查是否有真正錯位的空白列 (排除 itemHeaderRowIdx, typeRowIdx, dateRowIdx)
  for (var rCheck = 0; rCheck < searchUpper; rCheck++) {
    if (rCheck === result.itemHeaderRowIdx || rCheck === result.typeRowIdx || rCheck === result.dateRowIdx) continue;
    var isBlank = true;
    for (var cK = 0; cK < data[rCheck].length; cK++) {
      if ((data[rCheck][cK] || '').toString().trim() !== '') {
        isBlank = false;
        break;
      }
    }
    if (isBlank) {
      result.misplacedRowIndices.push(rCheck);
    }
  }

  if (result.firstStudentRowIdx === -1) {
    var maxTopRow = Math.max(result.itemHeaderRowIdx, result.typeRowIdx, result.dateRowIdx);
    result.firstStudentRowIdx = maxTopRow !== -1 ? (maxTopRow + 1) : 3;
  }

  return result;
}

/**
 * 確保平時成績分頁在「項目列」與「學生成績列」之間擁有「作業／考試」類別列
 * （若已有且無缺漏則 0ms 直接返回，不執行任何試算表寫入；若需建立或清理錯位列則透過 LockService 安全批次寫入）
 */
function ensureAssessmentTypeRow_(sheet, optData) {
  var data = optData || null;
  try {
    if (!data) {
      data = sheet.getDataRange().getValues();
    }
    if (!data || data.length < 2) return data;

    var loc = locateSubjectSheetRows_(data);
    if (loc.itemHeaderRowIdx === -1) return data;

    var itemHeaderRow = data[loc.itemHeaderRowIdx];
    var lastActiveCol = Math.max(loc.seatCol, loc.nameCol);
    for (var c = 0; c < itemHeaderRow.length; c++) {
      if ((itemHeaderRow[c] || '').toString().trim() !== '') {
        lastActiveCol = c;
      }
    }

    var hasMisplaced = loc.misplacedRowIndices.length > 0;
    var needsInsert = loc.typeRowIdx === -1;
    var needsFillMissingCols = false;

    if (!needsInsert && loc.typeRowIdx < data.length) {
      var existingTypeRow = data[loc.typeRowIdx];
      for (var checkCol = 0; checkCol <= lastActiveCol; checkCol++) {
        if (checkCol === loc.seatCol || checkCol === loc.nameCol) continue;
        var titleText = (itemHeaderRow[checkCol] || '').toString().trim();
        var curTypeText = (existingTypeRow[checkCol] || '').toString().trim();
        if (titleText && !curTypeText) {
          needsFillMissingCols = true;
          break;
        }
      }
    }

    // 快速通道：若無錯位列、且已有完整「作業／考試」列，直接回傳記憶體資料（零寫入 API 呼叫）
    if (!hasMisplaced && !needsInsert && !needsFillMissingCols) {
      return data;
    }

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(800)) {
      return data;
    }

    try {
      data = sheet.getDataRange().getValues();
      loc = locateSubjectSheetRows_(data);
      if (loc.itemHeaderRowIdx === -1) return data;

      // Step 1: 清理任何誤插在項目列上方的空白列或重複類別列（由下往上刪除）
      if (loc.misplacedRowIndices.length > 0) {
        var toDelete = loc.misplacedRowIndices.slice().sort(function(a, b) { return b - a; });
        for (var d = 0; d < toDelete.length; d++) {
          sheet.deleteRow(toDelete[d] + 1);
        }
        data = sheet.getDataRange().getValues();
        loc = locateSubjectSheetRows_(data);
        if (loc.itemHeaderRowIdx === -1) return data;
      }

      itemHeaderRow = data[loc.itemHeaderRowIdx];
      lastActiveCol = Math.max(loc.seatCol, loc.nameCol);
      for (var c2 = 0; c2 < itemHeaderRow.length; c2++) {
        if ((itemHeaderRow[c2] || '').toString().trim() !== '') {
          lastActiveCol = c2;
        }
      }
      var numCols = lastActiveCol + 1;

      if (loc.typeRowIdx === -1) {
        // Step 2: 在真正的項目列 (itemHeaderRowIdx) 下方插入「作業／考試」列
        sheet.insertRowAfter(loc.itemHeaderRowIdx + 1);
        var typeRow1Idx = loc.itemHeaderRowIdx + 2; // 1-indexed
        var typeValues = [];
        var bgColors = [];
        var fontColors = [];

        for (var col = 0; col < numCols; col++) {
          if (col === loc.seatCol) {
            typeValues.push('評量');
            bgColors.push('#475569');
            fontColors.push('#ffffff');
          } else if (col === loc.nameCol) {
            typeValues.push('作業/考試');
            bgColors.push('#475569');
            fontColors.push('#ffffff');
          } else {
            var rawTitle = (itemHeaderRow[col] || '').toString().trim();
            var autoType = inferDefaultAssessmentType(rawTitle);
            typeValues.push(autoType);
            if (autoType === '作業') {
              bgColors.push('#fef3c7');
              fontColors.push('#92400e');
            } else if (autoType === '考試') {
              bgColors.push('#e0f2fe');
              fontColors.push('#075985');
            } else {
              bgColors.push('#f5f5f4');
              fontColors.push('#78716c');
            }
          }
        }

        sheet.getRange(typeRow1Idx, 1, 1, numCols)
          .setValues([typeValues])
          .setBackgrounds([bgColors])
          .setFontColors([fontColors])
          .setFontWeight('bold')
          .setFontSize(10)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        sheet.setRowHeight(typeRow1Idx, 26);

        if (numCols > 2) {
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(['作業', '考試', '不計分'], true)
            .setAllowInvalid(true)
            .build();
          sheet.getRange(typeRow1Idx, 3, 1, numCols - 2).setDataValidation(rule);
        }
        data = sheet.getDataRange().getValues();
      } else {
        // Step 3: 已有「作業／考試」列 -> 批次補齊新增欄位之預設類別（100% 保留老師手動修改值）
        var typeRow1IdxExisting = loc.typeRowIdx + 1;
        var curRow = data[loc.typeRowIdx];
        var updatedValues = [];
        var hasAnyUpdate = false;
        for (var colIdx = 0; colIdx < numCols; colIdx++) {
          var curVal = (curRow[colIdx] !== undefined && curRow[colIdx] !== null) ? curRow[colIdx].toString().trim() : '';
          if (colIdx === loc.seatCol) {
            updatedValues.push(curVal || '評量');
          } else if (colIdx === loc.nameCol) {
            updatedValues.push(curVal || '作業/考試');
          } else {
            var itemTitle = (itemHeaderRow[colIdx] || '').toString().trim();
            if (itemTitle && !curVal) {
              var inferred = inferDefaultAssessmentType(itemTitle);
              updatedValues.push(inferred);
              hasAnyUpdate = true;
            } else {
              updatedValues.push(curVal);
            }
          }
        }
        if (hasAnyUpdate) {
          sheet.getRange(typeRow1IdxExisting, 1, 1, numCols).setValues([updatedValues]);
          if (loc.dateRowIdx === 0) {
            for (var cFix = 3; cFix <= numCols; cFix++) {
              setupSchemeCColumnHeader_(sheet, cFix);
            }
          }
          data = sheet.getDataRange().getValues();
        }
      }
    } finally {
      try { lock.releaseLock(); } catch (ignoreLock) {}
    }
  } catch (ignore) {}
  return data;
}

/**
 * 結合「項目名稱」與老師在試算表設定的「作業／考試」類別，判斷平時評量所屬分項
 */
function classifyScoreCategory(itemName, manualAssessmentType) {
  var h = (itemName || '').toString().trim();
  var aType = (manualAssessmentType || '').toString().trim() || inferDefaultAssessmentType(h);

  // 若老師設定為「作業」（或預設判為作業）
  if (aType === '作業' || aType.indexOf('作業') !== -1 || aType.indexOf('習') !== -1) {
    if (h.indexOf('習') !== -1 || aType.indexOf('習') !== -1) {
      return { key: 'workbook', label: '習作平時成績', shortLabel: '習作', icon: '📘', assessmentType: aType || '作業' };
    }
    return { key: 'homework', label: '作業卷平時成績', shortLabel: '作業卷', icon: '📝', assessmentType: aType || '作業' };
  }

  // 若老師設定為「考試」（或預設判為考試）
  if (aType === '考試' || aType.indexOf('考') !== -1 || aType.indexOf('測') !== -1) {
    if (h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1 || h.indexOf('生字') !== -1 || h.indexOf('單字') !== -1) {
      return { key: 'foundation', label: '語詞與基礎平時', shortLabel: '語詞/基礎', icon: '🔤', assessmentType: aType || '考試' };
    }
    return { key: 'quiz', label: '單元測驗卷成績', shortLabel: '測驗卷', icon: '🎯', assessmentType: aType || '考試' };
  }

  // 其他自訂標籤則依項目名稱回退判斷，並保留老師自訂的 assessmentType 標籤顯示於網站上
  if (h.indexOf('習') !== -1) {
    return { key: 'workbook', label: '習作平時成績', shortLabel: '習作', icon: '📘', assessmentType: aType };
  }
  if (h.indexOf('作業') !== -1 || h.indexOf('練習') !== -1 || h.indexOf('學習單') !== -1) {
    return { key: 'homework', label: '作業卷平時成績', shortLabel: '作業卷', icon: '📝', assessmentType: aType };
  }
  if (h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1) {
    return { key: 'foundation', label: '語詞與基礎平時', shortLabel: '語詞/基礎', icon: '🔤', assessmentType: aType };
  }
  return { key: 'quiz', label: '單元測驗卷成績', shortLabel: '測驗卷', icon: '🎯', assessmentType: aType };
}

/**
 * 從項目名稱中提取「單元標籤」(支援範圍單元如 L1-2, L1-3, L4-5, L1-6，以及單一單元 L1, L2, 第3課 -> L3)
 */
function extractUnitCode(itemName) {
  var h = (itemName || '').toString().trim();
  var mRange = h.match(/[LUlu]\s*(\d{1,2}\s*[-~～]\s*\d{1,2})/);
  if (mRange) {
    return 'L' + mRange[1].replace(/\s+/g, '').replace(/[~～]/g, '-');
  }

  var m1 = h.match(/[LUlu]\s*(\d{1,2})/);
  if (m1) return 'L' + parseInt(m1[1], 10);

  var m2 = h.match(/第([一二三四五六七八九十\d]+)[課單元回]/);
  if (m2) {
    var chMap = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
    var v = chMap[m2[1]] || parseInt(m2[1], 10);
    if (!isNaN(v)) return 'L' + v;
  }

  var m3 = h.match(/(\d{1,2}\s*[-~～]\s*\d{1,2})/);
  if (m3) return 'L' + m3[1].replace(/\s+/g, '').replace(/[~～]/g, '-');

  var m4 = h.match(/(\d{1,2})$/);
  if (m4) return 'L' + parseInt(m4[1], 10);

  return '平時';
}

/**
 * 依分數取得平時表現等第標籤與色系
 */
function getScoreGradeInfo(score) {
  if (score === null || score === undefined || score === '' || isNaN(Number(score))) {
    return { level: 'none', label: '尚無成績', badgeColor: 'stone' };
  }
  var n = Number(score);
  if (n >= 90) return { level: 'excellent', label: '🌟 精熟優異', badgeColor: 'emerald' };
  if (n >= 80) return { level: 'good', label: '⭐ 穩健良好', badgeColor: 'amber' };
  if (n >= 70) return { level: 'basic', label: '🌿 基礎達標', badgeColor: 'sky' };
  return { level: 'needs_review', label: '🌱 待加強複習', badgeColor: 'rose' };
}

/**
 * 辨識成績格中的請假紀錄：
 * 公 / 公假 -> 公假 (🏛️)
 * 事 / 事假 -> 事假 (📋)
 * 病 / 病假 -> 病假 (💊)
 * 喪 / 喪假 -> 喪假 (🕯️)
 */
function detectLeaveType(val) {
  if (val === null || val === undefined) return null;
  var str = val.toString().trim().replace(/[\s\u3000]+/g, '');
  if (!str) return null;

  if (str === '公' || str === '公假' || /^公假?$/i.test(str)) {
    return { type: '公假', short: '公', icon: '🏛️', color: 'sky' };
  }
  if (str === '事' || str === '事假' || /^事假?$/i.test(str)) {
    return { type: '事假', short: '事', icon: '📋', color: 'amber' };
  }
  if (str === '病' || str === '病假' || /^病假?$/i.test(str)) {
    return { type: '病假', short: '病', icon: '💊', color: 'rose' };
  }
  if (str === '喪' || str === '喪假' || /^喪假?$/i.test(str)) {
    return { type: '喪假', short: '喪', icon: '🕯️', color: 'slate' };
  }
  if (str === '假' || str === '請假' || /^請?假$/i.test(str)) {
    return { type: '請假', short: '假', icon: '📋', color: 'amber' };
  }

  // 容錯包含型判斷（例如 "公（假）"、"事假未考" 等）
  if (str.indexOf('公') !== -1 && (str.indexOf('假') !== -1 || str.length === 1)) {
    return { type: '公假', short: '公', icon: '🏛️', color: 'sky' };
  }
  if (str.indexOf('事') !== -1 && (str.indexOf('假') !== -1 || str.length === 1)) {
    return { type: '事假', short: '事', icon: '📋', color: 'amber' };
  }
  if (str.indexOf('病') !== -1 && (str.indexOf('假') !== -1 || str.length === 1)) {
    return { type: '病假', short: '病', icon: '💊', color: 'rose' };
  }
  if (str.indexOf('喪') !== -1 && (str.indexOf('假') !== -1 || str.length === 1)) {
    return { type: '喪假', short: '喪', icon: '🕯️', color: 'slate' };
  }
  if (str.indexOf('假') !== -1) {
    return { type: '請假', short: '假', icon: '📋', color: 'amber' };
  }

  return null;
}

/**
 * 計算考卷難易度指數（Test Difficulty Index, P*，方法三抗極端值校正）
 * 公式：
 *   平均得分率 P_mean = X_bar / 100
 *   中位數得分率 P_med = Median / 100
 *   遇偏態（|X_bar - Median| > 3）時採抗極端值加權：P* = 0.6 * P_mean + 0.4 * P_med
 *   無偏態（|X_bar - Median| <= 3）時：P* = P_mean
 * 分級標準（以 P* 換算百分制分數 indexScore 判定）：
 *   >= 80 分：🟢 偏易
 *   65 ~ 79.9 分：🎯 適中
 *   50 ~ 64.9 分：🔥 偏難
 *   < 50 分：⚡ 極具挑戰
 * 強制附註：*難易度以班級成績估計。
 */
function computeTestDifficultyIndex_(scores) {
  if (!scores || scores.length === 0) return null;
  var valid = [];
  var sum = 0;
  for (var i = 0; i < scores.length; i++) {
    var v = Number(scores[i]);
    if (!isNaN(v)) {
      valid.push(v);
      sum += v;
    }
  }
  var n = valid.length;
  if (n === 0) return null;

  valid.sort(function(a, b) { return a - b; });
  var rawMean = sum / n;
  var rawMedian = (n % 2 === 1)
    ? valid[Math.floor(n / 2)]
    : (valid[n / 2 - 1] + valid[n / 2]) / 2;

  var mean = Number(rawMean.toFixed(1));
  var median = Number(rawMedian.toFixed(1));
  var skewDiff = Number(Math.abs(rawMean - rawMedian).toFixed(1));
  var isSkewed = Math.abs(rawMean - rawMedian) > 3;

  var pMean = rawMean / 100;
  var pMed = rawMedian / 100;
  var pStarRaw = isSkewed ? (0.6 * pMean + 0.4 * pMed) : pMean;
  var indexScore = Number((pStarRaw * 100).toFixed(1));
  var pStar = Number(pStarRaw.toFixed(3));

  var level = 'moderate';
  var icon = '🎯';
  var shortLabel = '適中';
  var color = 'sky';

  if (indexScore >= 80) {
    level = 'easy';
    icon = '🟢';
    shortLabel = '偏易';
    color = 'emerald';
  } else if (indexScore >= 65) {
    level = 'moderate';
    icon = '🎯';
    shortLabel = '適中';
    color = 'sky';
  } else if (indexScore >= 50) {
    level = 'hard';
    icon = '🔥';
    shortLabel = '偏難';
    color = 'amber';
  } else {
    level = 'challenging';
    icon = '⚡';
    shortLabel = '極具挑戰';
    color = 'rose';
  }

  return {
    pStar: pStar,
    indexScore: indexScore,
    mean: mean,
    median: median,
    pMean: Number(pMean.toFixed(3)),
    pMed: Number(pMed.toFixed(3)),
    skewDiff: skewDiff,
    isSkewed: isSkewed,
    level: level,
    icon: icon,
    shortLabel: shortLabel,
    label: icon + ' ' + shortLabel,
    color: color,
    sampleCount: n,
    note: '*難易度以班級成績估計。'
  };
}

/**
 * 解析單一平時成績工作表（自動確保「作業／考試」列存在，並僅回傳該試算表實際存在的評量分項！）
 */
function parseDailySubjectSheet(sheet) {
  var rawData = sheet.getDataRange().getValues();
  if (!rawData || rawData.length < 2) return null;

  var data = ensureAssessmentTypeRow_(sheet, rawData);
  var loc = locateSubjectSheetRows_(data);
  if (loc.itemHeaderRowIdx === -1) return null;

  var sheetName = sheet.getName();
  var sheetTitle = sheetName;

  // 僅當第 1 列不是「登記日期列」，且為傳統標題列時，才嘗試自第 1 列讀取自訂大標題
  if (loc.itemHeaderRowIdx > 0 && loc.dateRowIdx !== 0) {
    var titleCandidates = data[0].filter(function(cell) {
      if (!cell || cell instanceof Date) return false;
      var s = cell.toString().trim();
      if (s.indexOf('GMT') !== -1 || s.indexOf('台北標準時間') !== -1 || s.indexOf('UTC') !== -1 || s.indexOf('+0800') !== -1) return false;
      if (/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(s) || /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(s)) return false;
      return s && s !== '📌 學生名單' && s !== '📌 登記日期' && s !== '座號' && s !== '姓名' && !extractDateFromTitle_(s);
    });
    if (titleCandidates.length > 0) sheetTitle = titleCandidates[0].toString().trim();
  }

  var itemHeaderRow = data[loc.itemHeaderRowIdx];
  var seatCol = loc.seatCol;
  var nameCol = loc.nameCol;
  var typeRow = (loc.typeRowIdx !== -1 && loc.typeRowIdx < data.length) ? data[loc.typeRowIdx] : null;
  var dateRow = (loc.dateRowIdx !== -1 && loc.dateRowIdx < data.length) ? data[loc.dateRowIdx] : (data[0] || null);
  var dataStartRowIdx = loc.firstStudentRowIdx;

  var itemCols = [];
  var sheetPillarColCounts = { workbook: 0, homework: 0, foundation: 0, quiz: 0 };

  for (var col = 0; col < itemHeaderRow.length; col++) {
    if (col === seatCol || col === nameCol) continue;

    var rawItemTitle = (itemHeaderRow[col] || '').toString().trim();
    if (!rawItemTitle) continue;

    var manualType = typeRow ? (typeRow[col] || '').toString().trim() : '';
    if (manualType === '不計分' || manualType === '略過' || manualType === '加分' || manualType === '扣分') {
      continue;
    }
    if (!manualType && isSummaryColumnHeader(rawItemTitle)) {
      continue;
    }

    var catInfo = classifyScoreCategory(rawItemTitle, manualType);
    var unitCode = extractUnitCode(rawItemTitle);
    sheetPillarColCounts[catInfo.key] = (sheetPillarColCounts[catInfo.key] || 0) + 1;

    var colDateVal = (dateRow && dateRow[col] !== undefined && dateRow[col] !== null) ? dateRow[col] : null;

    itemCols.push({
      colIdx: col,
      title: rawItemTitle,
      shortTitle: rawItemTitle,
      assessmentType: catInfo.assessmentType || '作業',
      pillarKey: catInfo.key,
      pillarLabel: catInfo.label,
      pillarShort: catInfo.shortLabel,
      pillarIcon: catInfo.icon,
      unitCode: unitCode,
      date: colDateVal
    });
  }

  var rawStudents = [];
  var colSums = {};
  var colCounts = {};
  var colScoresList = {};

  for (var i = 0; i < itemCols.length; i++) {
    colSums[itemCols[i].colIdx] = 0;
    colCounts[itemCols[i].colIdx] = 0;
    colScoresList[itemCols[i].colIdx] = [];
  }

  for (var rIdx = dataStartRowIdx; rIdx < data.length; rIdx++) {
    var row = data[rIdx];
    if (isAssessmentTypeRow_(row, seatCol, nameCol)) continue;
    var seatVal = row[seatCol] !== undefined ? row[seatCol].toString().trim() : '';
    var nameVal = row[nameCol] !== undefined ? row[nameCol].toString().trim() : '';

    // 必須同時具備座號與姓名才視為有效學生列（自動略過表尾無姓名的班均或統計列）
    if (!seatVal || !nameVal) continue;
    if (seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1 || nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal);
    rawStudents.push({
      rowIdx: rIdx,
      seat: normSeat,
      name: nameVal,
      row: row
    });

    for (var j = 0; j < itemCols.length; j++) {
      var ic = itemCols[j];
      var cellVal = row[ic.colIdx];
      var isLeaveCell = detectLeaveType(cellVal) !== null;
      var num = (!isLeaveCell && cellVal !== '' && cellVal !== null && !isNaN(Number(cellVal))) ? Number(cellVal) : null;
      if (num !== null) {
        colSums[ic.colIdx] += num;
        colCounts[ic.colIdx] += 1;
        colScoresList[ic.colIdx].push(num);
      }
    }
  }

  var colClassAvgs = {};
  var colDifficulties = {};
  var subjectAssessments = [];
  for (var k = 0; k < itemCols.length; k++) {
    var cIdx = itemCols[k].colIdx;
    colClassAvgs[cIdx] = colCounts[cIdx] > 0 ? Number((colSums[cIdx] / colCounts[cIdx]).toFixed(1)) : null;
    colDifficulties[cIdx] = computeTestDifficultyIndex_(colScoresList[cIdx]);
    if (colCounts[cIdx] > 0 && colDifficulties[cIdx]) {
      subjectAssessments.push({
        colIdx: itemCols[k].colIdx + 1,
        name: itemCols[k].title,
        date: itemCols[k].date,
        unit: itemCols[k].unitCode || '平時',
        assessmentType: itemCols[k].assessmentType,
        category: itemCols[k].pillarKey,
        categoryLabel: itemCols[k].pillarShort,
        categoryIcon: itemCols[k].pillarIcon,
        classAvg: colClassAvgs[cIdx],
        median: colDifficulties[cIdx].median,
        difficulty: colDifficulties[cIdx]
      });
    }
  }

  var studentMap = {};
  var classOverallSum = 0;
  var classOverallCount = 0;
  var allStudentAverages = [];
  var pillarClassSums = { workbook: 0, homework: 0, foundation: 0, quiz: 0 };
  var pillarClassCounts = { workbook: 0, homework: 0, foundation: 0, quiz: 0 };

  for (var s = 0; s < rawStudents.length; s++) {
    var st = rawStudents[s];
    var items = [];
    var pillarBuckets = {
      workbook: { scores: [] },
      homework: { scores: [] },
      foundation: { scores: [] },
      quiz: { scores: [] }
    };

    var allValidScores = [];
    var unitMap = {};
    var unitOrder = [];

    for (var m = 0; m < itemCols.length; m++) {
      var colMeta = itemCols[m];
      var rawVal = st.row[colMeta.colIdx];
      var leaveInfo = detectLeaveType(rawVal);
      var isLeave = leaveInfo !== null;
      var hasScoresInCol = colCounts[colMeta.colIdx] > 0;

      // 若整欄無人有成績且該生亦無成績也無請假，則略過
      if (!hasScoresInCol && !isLeave && (rawVal === '' || rawVal === null)) {
        continue;
      }

      var numScore = (!isLeave && rawVal !== '' && rawVal !== null && !isNaN(Number(rawVal))) ? Number(rawVal) : null;
      var classAvg = colClassAvgs[colMeta.colIdx];
      var gradeInfo = isLeave
        ? { level: 'leave', label: leaveInfo.icon + ' ' + leaveInfo.type, badgeColor: leaveInfo.color }
        : getScoreGradeInfo(numScore);
      var diffVal = (numScore !== null && classAvg !== null) ? Number((numScore - classAvg).toFixed(1)) : null;

      var bkt = pillarBuckets[colMeta.pillarKey] || pillarBuckets.quiz;
      if (numScore !== null) {
        bkt.scores.push(numScore);
        allValidScores.push(numScore);
      } else if (isLeave) {
        bkt.leaves = bkt.leaves || [];
        bkt.leaves.push(leaveInfo.type);
      }

      var uKey = colMeta.unitCode || '平時';
      if (!unitMap[uKey]) {
        unitMap[uKey] = {
          unit: uKey,
          scores: [],
          classAvgs: [],
          byPillar: { workbook: [], homework: [], foundation: [], quiz: [] },
          classByPillar: { workbook: [], homework: [], foundation: [], quiz: [] }
        };
        unitOrder.push(uKey);
      }
      if (numScore !== null) {
        unitMap[uKey].scores.push(numScore);
        unitMap[uKey].byPillar[colMeta.pillarKey].push(numScore);
      }
      if (classAvg !== null) {
        unitMap[uKey].classAvgs.push(classAvg);
        unitMap[uKey].classByPillar[colMeta.pillarKey].push(classAvg);
      }

      var diffInfo = colDifficulties[colMeta.colIdx] || null;
      items.push({
        name: colMeta.title,
        title: colMeta.title,
        shortTitle: colMeta.shortTitle,
        unit: uKey,
        unitCode: uKey,
        assessmentType: colMeta.assessmentType,
        category: colMeta.pillarKey,
        pillarKey: colMeta.pillarKey,
        categoryLabel: colMeta.pillarShort,
        pillarLabel: colMeta.pillarLabel,
        pillarShort: colMeta.pillarShort,
        categoryIcon: colMeta.pillarIcon,
        score: numScore,
        rawScore: (rawVal !== null && rawVal !== undefined) ? rawVal.toString().trim() : '',
        isLeave: isLeave,
        leaveType: isLeave ? leaveInfo.type : null,
        leaveShort: isLeave ? leaveInfo.short : null,
        leaveIcon: isLeave ? leaveInfo.icon : null,
        classAvg: classAvg,
        median: diffInfo ? diffInfo.median : null,
        difficulty: diffInfo,
        diffFromClass: diffVal,
        gradeLevel: gradeInfo.level,
        gradeLabel: gradeInfo.label,
        badgeColor: gradeInfo.badgeColor
      });
    }

    var pillarKeys = ['workbook', 'homework', 'foundation', 'quiz'];
    var pillarMetaMap = {
      workbook: { label: '習作平時成績', shortLabel: '習作平均', icon: '📘', assessmentType: '作業' },
      homework: { label: '作業卷平時成績', shortLabel: '作業卷平均', icon: '📝', assessmentType: '作業' },
      foundation: { label: '語詞與基礎平時', shortLabel: '語詞/基礎均', icon: '🔤', assessmentType: '考試' },
      quiz: { label: '單元測驗卷成績', shortLabel: '測驗卷平均', icon: '🎯', assessmentType: '考試' }
    };

    var pillarsArray = [];
    var pillarsSummaryObj = {};

    for (var pkIdx = 0; pkIdx < pillarKeys.length; pkIdx++) {
      var pk = pillarKeys[pkIdx];
      // 若該試算表分頁根本沒有此類別的欄位（例如數學科沒有語詞測驗或作業卷），則不產生該分項卡片！
      if (!sheetPillarColCounts[pk]) continue;

      var pb = pillarBuckets[pk];
      var pAvg = null;
      var pMax = null;
      var pMin = null;
      if (pb.scores.length > 0) {
        var pSum = 0;
        for (var ps = 0; ps < pb.scores.length; ps++) {
          var v = pb.scores[ps];
          pSum += v;
          if (pMax === null || v > pMax) pMax = v;
          if (pMin === null || v < pMin) pMin = v;
        }
        pAvg = Number((pSum / pb.scores.length).toFixed(1));
        pillarClassSums[pk] += pAvg;
        pillarClassCounts[pk] += 1;
      }
      var pGrade = getScoreGradeInfo(pAvg);
      var pObj = {
        key: pk,
        label: pillarMetaMap[pk].label,
        shortLabel: pillarMetaMap[pk].shortLabel,
        icon: pillarMetaMap[pk].icon,
        assessmentType: pillarMetaMap[pk].assessmentType,
        average: pAvg,
        maxScore: pMax,
        minScore: pMin,
        count: pb.scores.length,
        leaveCount: (pb.leaves ? pb.leaves.length : 0),
        badge: pGrade.label,
        badgeColor: pGrade.badgeColor
      };
      pillarsArray.push(pObj);
      pillarsSummaryObj[pk] = pObj;
    }

    var overallAvg = null;
    if (allValidScores.length > 0) {
      var totalS = 0;
      for (var as = 0; as < allValidScores.length; as++) totalS += allValidScores[as];
      overallAvg = Number((totalS / allValidScores.length).toFixed(1));
      classOverallSum += overallAvg;
      classOverallCount += 1;
      allStudentAverages.push(overallAvg);
    }

    var avgHelper = function(arr) {
      if (!arr || arr.length === 0) return null;
      var sum = 0;
      for (var idx = 0; idx < arr.length; idx++) sum += arr[idx];
      return Number((sum / arr.length).toFixed(1));
    };

    var unitTimeline = unitOrder.map(function(u) {
      var uObj = unitMap[u];
      var uAvg = avgHelper(uObj.scores);
      var cAvg = avgHelper(uObj.classAvgs);
      var wbAvg = avgHelper(uObj.byPillar.workbook);
      var wbClass = avgHelper(uObj.classByPillar.workbook);
      var hwAvg = avgHelper(uObj.byPillar.homework);
      var hwClass = avgHelper(uObj.classByPillar.homework);
      var fdAvg = avgHelper(uObj.byPillar.foundation);
      var fdClass = avgHelper(uObj.classByPillar.foundation);
      var qzAvg = avgHelper(uObj.byPillar.quiz);
      var qzClass = avgHelper(uObj.classByPillar.quiz);

      return {
        unit: u,
        overallScore: uAvg,
        classOverallAvg: cAvg,
        classAvg: cAvg,
        workbook: wbAvg,
        classWorkbook: wbClass !== null ? wbClass : cAvg,
        workbookClassAvg: wbClass !== null ? wbClass : cAvg,
        homework: hwAvg,
        classHomework: hwClass !== null ? hwClass : cAvg,
        homeworkClassAvg: hwClass !== null ? hwClass : cAvg,
        foundation: fdAvg,
        classFoundation: fdClass !== null ? fdClass : cAvg,
        foundationClassAvg: fdClass !== null ? fdClass : cAvg,
        quiz: qzAvg,
        classQuiz: qzClass !== null ? qzClass : cAvg,
        quizClassAvg: qzClass !== null ? qzClass : cAvg
      };
    });

    var shortSubj = inferSubjectShortName(sheetName, sheetTitle);
    var leaveItems = items.filter(function(it) { return it.isLeave; });
    var stDataObj = {
      subject: shortSubj,
      sheetName: sheetName,
      subjectTitle: sheetTitle,
      subjectShortName: shortSubj,
      overallAverage: overallAvg,
      totalItems: allValidScores.length,
      totalScoresCount: allValidScores.length,
      leaveCount: leaveItems.length,
      leaves: leaveItems.map(function(it) { return { name: it.name, type: it.leaveType, icon: it.leaveIcon }; }),
      pillars: pillarsArray,
      pillarsMap: pillarsSummaryObj,
      unitTimeline: unitTimeline,
      assessmentsDifficulty: subjectAssessments,
      items: items
    };
    studentMap[st.seat] = stDataObj;
    if (st.name) {
      studentMap['NAME_' + st.name] = stDataObj;
    }
  }

  var subjectClassAvg = classOverallCount > 0 ? Number((classOverallSum / classOverallCount).toFixed(1)) : null;
  var pillarClassAvgs = {};
  ['workbook', 'homework', 'foundation', 'quiz'].forEach(function(pk) {
    pillarClassAvgs[pk] = pillarClassCounts[pk] > 0 ? Number((pillarClassSums[pk] / pillarClassCounts[pk]).toFixed(1)) : null;
  });

  Object.keys(studentMap).forEach(function(seatKey) {
    if (seatKey.indexOf('NAME_') === 0) return;
    var stRecord = studentMap[seatKey];
    stRecord.classOverallAverage = subjectClassAvg;
    stRecord.diffFromClassAvg = (stRecord.overallAverage !== null && subjectClassAvg !== null) ? Number((stRecord.overallAverage - subjectClassAvg).toFixed(1)) : null;

    if (stRecord.overallAverage !== null && allStudentAverages.length > 0) {
      var belowCnt = 0;
      for (var idx = 0; idx < allStudentAverages.length; idx++) {
        if (allStudentAverages[idx] < stRecord.overallAverage) belowCnt++;
      }
      stRecord.prValue = Math.max(1, Math.min(99, Math.round((belowCnt / allStudentAverages.length) * 100)));
    } else {
      stRecord.prValue = null;
    }

    stRecord.pillars.forEach(function(pItem) {
      var cAvg = pillarClassAvgs[pItem.key];
      pItem.classAverage = cAvg;
      pItem.diffFromClass = (pItem.average !== null && cAvg !== null) ? Number((pItem.average - cAvg).toFixed(1)) : null;
    });
    stRecord.teacherAdvice = buildDailyScoreDiagnosticAdvice(stRecord);
  });

  return {
    sheetName: sheetName,
    subjectTitle: sheetTitle,
    subjectShortName: inferSubjectShortName(sheetName, sheetTitle),
    classOverallAverage: subjectClassAvg,
    assessments: subjectAssessments,
    itemCols: itemCols,
    studentMap: studentMap
  };
}

/**
 * 保留教師輸入之工作表分頁名稱作為科目顯示名稱（教師輸入什麼科目名稱，網站就顯示什麼科目名稱！）
 * 內建嚴密防護，徹底杜絕 GMT、台北標準時間、UTC 或日期字串洩漏作為科目名稱
 */
function inferSubjectShortName(sheetName, sheetTitle) {
  var isDateOrTimezoneString = function(s) {
    if (!s) return true;
    var str = s.toString().trim();
    if (!str) return true;
    if (str.indexOf('GMT') !== -1 || str.indexOf('台北標準時間') !== -1 || str.indexOf('UTC') !== -1 || str.indexOf('+0800') !== -1) return true;
    if (/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(str)) return true;
    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(str)) return true;
    return false;
  };

  var exactName = (sheetName || '').toString().trim();
  if (exactName && !isDateOrTimezoneString(exactName)) return exactName;
  var title = (sheetTitle || '').toString().trim();
  if (title && !isDateOrTimezoneString(title)) {
    return title;
  }
  return '平時成績';
}

/**
 * 產生個人化平時成績診斷分析評語
 */
function buildDailyScoreDiagnosticAdvice(stRecord) {
  var sub = stRecord.subjectShortName || '本科';
  var parts = [];

  var pMap = stRecord.pillarsMap || {};
  var wb = pMap.workbook ? pMap.workbook.average : null;
  var qz = pMap.quiz ? pMap.quiz.average : null;

  if (stRecord.overallAverage !== null) {
    var classAvgSuffix = (stRecord.classOverallAverage !== null && stRecord.classOverallAverage !== undefined)
      ? '（班均 ' + stRecord.classOverallAverage + ' 分）'
      : '';
    if (stRecord.overallAverage >= 90) {
      parts.push('【' + sub + '】平時總平均達 ' + stRecord.overallAverage + ' 分' + classAvgSuffix + '，平時表現非常扎實穩健！');
    } else if (stRecord.overallAverage >= 80) {
      parts.push('【' + sub + '】平時總平均為 ' + stRecord.overallAverage + ' 分' + classAvgSuffix + '，各類平時表現穩定良好。');
    } else {
      parts.push('【' + sub + '】平時總平均為 ' + stRecord.overallAverage + ' 分' + classAvgSuffix + '，建議透過各單元平時測驗的分數變化，逐步鞏固核心概念。');
    }
  }

  if (wb !== null && qz !== null && (wb - qz >= 8)) {
    parts.push('從分項平均觀察，習作平時平均（' + wb + '分）明顯高於單元測驗卷平均（' + qz + '分），代表平時隨堂練習認真，但面對綜合測驗卷變化題型時較易失分，考前建議多練習綜合應用題。');
  } else if (qz !== null && qz >= 90) {
    parts.push('單元測驗卷平均達 ' + qz + ' 分，顯示各課核心觀念融會貫通，應試穩定度極佳！');
  }

  if (stRecord.unitTimeline && stRecord.unitTimeline.length >= 2) {
    var lowestUnit = null;
    var highestUnit = null;
    stRecord.unitTimeline.forEach(function(u) {
      var val = u.overallScore !== null ? u.overallScore : u.quiz;
      if (val !== null) {
        if (!lowestUnit || val < lowestUnit.val) lowestUnit = { unit: u.unit, val: val };
        if (!highestUnit || val > highestUnit.val) highestUnit = { unit: u.unit, val: val };
      }
    });
    if (highestUnit && lowestUnit && highestUnit.unit !== lowestUnit.unit && (highestUnit.val - lowestUnit.val >= 8)) {
      parts.push('在單元歷程走勢中，' + highestUnit.unit + '（均分 ' + highestUnit.val + '分）表現最佳；而 ' + lowestUnit.unit + '（均分 ' + lowestUnit.val + '分）相對偏低，建議期中/期末考前優先複習 ' + lowestUnit.unit + '！');
    }
  }

  if (stRecord.items && stRecord.items.length > 0) {
    var hardTestsWellDone = [];
    for (var hi = 0; hi < stRecord.items.length; hi++) {
      var it = stRecord.items[hi];
      if (!it.isLeave && it.score !== null && it.difficulty && (it.difficulty.level === 'hard' || it.difficulty.level === 'challenging')) {
        if (it.classAvg !== null && it.score >= it.classAvg) {
          hardTestsWellDone.push(it.name + '（難度' + it.difficulty.label + '，得分 ' + it.score + ' 分高於班均）');
        }
      }
    }
    if (hardTestsWellDone.length > 0) {
      parts.push('值得一提的是，在難度較高的評量如「' + hardTestsWellDone.slice(0, 2).join('、') + '」中仍能保持超越班均的表現，抗壓與融會貫通能力值得肯定！（*難易度以班級成績估計。）');
    }
  }

  if (stRecord.leaveCount > 0) {
    parts.push('（本科包含 ' + stRecord.leaveCount + ' 次請假紀錄，依校規皆免計分且不列入平均計算。）');
  }

  return parts.join(' ');
}

/**
 * 檢查系統連線狀態 (ping)
 */
function checkSystemStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  var sheets = ss.getSheets();
  var subjectNames = [];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    if (!isSystemSheet(sName)) {
      subjectNames.push(sName);
    }
  }

  if (!sSheet) {
    return {
      status: 'error',
      success: false,
      message: '找不到【學生名冊】分頁！請在試算表上方點選【⚡ MPScore-Daily 平時成績工具 > 🛠️ 建立初始表格】。'
    };
  }

  var studentCount = Math.max(0, sSheet.getLastRow() - 1);
  return {
    status: 'success',
    success: true,
    message: '連線成功！已載入 ' + studentCount + ' 位學生名冊與 ' + subjectNames.length + ' 個平時成績分頁。',
    spreadsheetName: ss.getName(),
    studentCount: studentCount,
    subjectList: subjectNames,
    authRequired: true,
    isCustomPinSet: getTeacherAdminKey_() !== '8888'
  };
}

/**
 * 家長端依 Token 查詢該生所有平時成績歷程資料 (高資安隔離)
 */
function getStudentDailyPortfolioByToken(token, clientDevice, skipCacheOpt, isSimulated) {
  var cleanToken = (token || '').toString().trim().toUpperCase();
  if (!cleanToken) {
    return { status: 'error', success: false, message: '請掃描有效的專屬 QR Code！' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { status: 'error', success: false, message: '系統尚未初始化【學生名冊】分頁。' };
  }

  var sData = sSheet.getDataRange().getValues();
  var matchedStudent = null;
  var matchedRowIdx = -1;

  for (var r = 1; r < sData.length; r++) {
    var rowTok = (sData[r][3] || '').toString().trim().toUpperCase();
    var enabled = sData[r][4] !== false && sData[r][4].toString().toUpperCase() !== 'FALSE';
    if (rowTok === cleanToken && enabled) {
      var cls = (sData[r][0] || '本班').toString().trim();
      var seat = normalizeSeatNumber(sData[r][1]);
      matchedStudent = {
        className: cls,
        seatNo: seat,
        seatNumber: seat,
        name: (sData[r][2] || '').toString().trim(),
        token: (sData[r][3] || '').toString().trim(),
        studentId: cls + '-' + seat
      };
      matchedRowIdx = r + 1; // 1-based row index in spreadsheet
      break;
    }
  }

  if (!matchedStudent) {
    return { status: 'error', success: false, message: '查無此專屬代碼或尚未啟用，請確認通知單上的 QR Code 是否正確。' };
  }

  // 記錄家長閱覽日誌（含 3 分鐘防抖保護。若為教師端模擬預覽模式則絕不記錄，保護真實數據純淨）
  if (!isSimulated) {
    try {
      recordParentAccessLog_(ss, sSheet, matchedStudent, matchedRowIdx, clientDevice);
    } catch (logErr) {}
  }

  var skipCache = Boolean(skipCacheOpt);
  var cache = CacheService.getScriptCache();
  var cacheKey = 'mpdaily_score_v8_' + cleanToken;
  var cachedJson = !skipCache ? cache.get(cacheKey) : null;
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (ignore) {}
  }

  var sheets = ss.getSheets();
  var subjectsData = [];
  var combinedAdvice = [];

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (parsed) {
      var stSubj = parsed.studentMap[matchedStudent.seatNumber] || parsed.studentMap['NAME_' + matchedStudent.name];
      if (stSubj) {
        subjectsData.push(stSubj);
        if (stSubj.teacherAdvice) combinedAdvice.push(stSubj.teacherAdvice);
      }
    }
  }

  matchedStudent.subjects = subjectsData;
  matchedStudent.teacherAdvice = combinedAdvice.join(' ');

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
  var responseObj = {
    status: 'success',
    success: true,
    student: matchedStudent,
    subjects: subjectsData,
    updatedAt: nowStr
  };

  try {
    cache.put(cacheKey, JSON.stringify(responseObj), 60);
  } catch (ignore) {}

  return responseObj;
}

/**
 * 記錄家長閱覽時間與累計次數（自動維護學生名冊與家長閱覽紀錄分頁，含 3 分鐘防抖保護）
 */
function recordParentAccessLog_(ss, sSheet, student, rowIdx, clientDevice) {
  var cache = CacheService.getScriptCache();
  var debounceKey = 'mpdaily_visit_debounce_' + student.token;
  var isDebounced = cache.get(debounceKey);

  var now = new Date();
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
  var device = (clientDevice || '').toString().trim() || '行動裝置';

  // 1. 確保【學生名冊】欄位 F（最後閱覽時間）與欄位 G（累積閱覽次數）之標題存在
  var lastCol = sSheet.getLastColumn();
  if (lastCol < 6) {
    sSheet.getRange(1, 6).setValue('最後閱覽時間')
      .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sSheet.setColumnWidth(6, 150);
  } else {
    var h6 = sSheet.getRange(1, 6).getValue();
    if (!h6 || h6.toString().indexOf('閱覽') === -1) {
      sSheet.getRange(1, 6).setValue('最後閱覽時間')
        .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      sSheet.setColumnWidth(6, 150);
    }
  }

  if (lastCol < 7) {
    sSheet.getRange(1, 7).setValue('累積閱覽次數')
      .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sSheet.setColumnWidth(7, 110);
  } else {
    var h7 = sSheet.getRange(1, 7).getValue();
    if (!h7 || h7.toString().indexOf('次數') === -1) {
      sSheet.getRange(1, 7).setValue('累積閱覽次數')
        .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      sSheet.setColumnWidth(7, 110);
    }
  }

  // 2. 若在 3 分鐘內重複開啟，只視為同一筆，不重複增加寫入與次數（防刷）
  if (isDebounced) {
    return;
  }
  cache.put(debounceKey, '1', 180);

  // 3. 更新【學生名冊】該生之最後閱覽時間與累積次數
  var currentCount = 0;
  try {
    var rawCount = sSheet.getRange(rowIdx, 7).getValue();
    currentCount = parseInt(rawCount || '0', 10);
    if (isNaN(currentCount)) currentCount = 0;
  } catch (e) {}

  sSheet.getRange(rowIdx, 6).setValue(nowStr).setHorizontalAlignment('center');
  sSheet.getRange(rowIdx, 7).setValue(currentCount + 1).setHorizontalAlignment('center');

  // 4. 寫入【家長閱覽紀錄】專屬日誌分頁
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  if (!logSheet) {
    logSheet = ss.insertSheet('家長閱覽紀錄');
    var logHeaders = [['閱覽時間', '班級', '座號', '姓名', '查詢代碼', '裝置環境']];
    logSheet.getRange(1, 1, 1, 6).setValues(logHeaders)
      .setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    try {
      logSheet.setFrozenRows(1);
    } catch (ignore) {}
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(2, 80);
    logSheet.setColumnWidth(3, 80);
    logSheet.setColumnWidth(4, 120);
    logSheet.setColumnWidth(5, 120);
    logSheet.setColumnWidth(6, 140);
  }

  logSheet.insertRowBefore(2);
  logSheet.getRange(2, 1, 1, 6).setValues([[
    nowStr,
    student.className || '本班',
    student.seatNo || '',
    student.name || '',
    student.token || '',
    device
  ]]).setHorizontalAlignment('center');

  // 維護紀錄上限，保留最新 1000 筆避免試算表肥大
  if (logSheet.getLastRow() > 1001) {
    try {
      logSheet.deleteRows(1002, logSheet.getLastRow() - 1001);
    } catch (ignore) {}
  }
}

/**
 * 讀取最新家長閱覽紀錄明細（供教師端即時儀表板與日誌明細視窗使用）
 */
function getRecentAccessLogs_(ss, maxCount) {
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  if (!logSheet || logSheet.getLastRow() <= 1) return [];
  var limit = Math.min(maxCount || 50, logSheet.getLastRow() - 1);
  var values = logSheet.getRange(2, 1, limit, 6).getValues();
  var logs = [];
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  for (var i = 0; i < values.length; i++) {
    var rawTime = values[i][0];
    var timeStr = '';
    if (rawTime instanceof Date) {
      timeStr = Utilities.formatDate(rawTime, tz, 'yyyy-MM-dd HH:mm');
    } else if (rawTime) {
      timeStr = rawTime.toString().trim();
    }
    logs.push({
      time: timeStr,
      className: (values[i][1] || '').toString().trim(),
      seatNo: (values[i][2] || '').toString().trim(),
      name: (values[i][3] || '').toString().trim(),
      token: (values[i][4] || '').toString().trim(),
      device: (values[i][5] || '').toString().trim()
    });
  }
  return logs;
}

/**
 * 讀取各學生專屬之近 10 次家長瀏覽時間紀錄對照表
 */
function getStudentAccessHistoryMap_(ss) {
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  if (!logSheet || logSheet.getLastRow() <= 1) return {};
  var numRows = Math.min(1000, logSheet.getLastRow() - 1);
  var data = logSheet.getRange(2, 1, numRows, 6).getValues();
  var map = {};
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';

  for (var i = 0; i < data.length; i++) {
    var rawTime = data[i][0];
    var timeStr = '';
    if (rawTime instanceof Date) {
      timeStr = Utilities.formatDate(rawTime, tz, 'yyyy-MM-dd HH:mm');
    } else if (rawTime) {
      timeStr = rawTime.toString().trim();
    }
    var tok = (data[i][4] || '').toString().trim().toUpperCase();
    var dev = (data[i][5] || '').toString().trim() || '行動裝置';
    if (!tok) continue;
    if (!map[tok]) map[tok] = [];
    if (map[tok].length < 10) {
      map[tok].push({ time: timeStr, device: dev });
    }
  }
  return map;
}

/**
 * 教師端同步名冊與全班各科平時成績統計（供 A4 通知小卡 PDF 預覽與全班總表使用）
 */
function getStudentListForTeacher() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { status: 'error', success: false, message: '找不到【學生名冊】分頁！請先建立初始表格。' };
  }

  var sData = sSheet.getDataRange().getValues();
  var students = [];
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var historyMap = getStudentAccessHistoryMap_(ss);

  for (var r = 1; r < sData.length; r++) {
    var cls = (sData[r][0] || '本班').toString().trim();
    var seat = normalizeSeatNumber(sData[r][1]);
    var name = (sData[r][2] || '').toString().trim();
    var token = (sData[r][3] || '').toString().trim();
    var lastViewedAt = '';
    var rawTime = sData[r][5];
    if (rawTime instanceof Date) {
      lastViewedAt = Utilities.formatDate(rawTime, tz, 'yyyy-MM-dd HH:mm');
    } else if (rawTime) {
      lastViewedAt = rawTime.toString().trim();
    }
    var rawCount = sData[r][6];
    var viewCount = parseInt(rawCount || '0', 10) || 0;
    if (!seat && !name) continue;

    var studentToken = token.toUpperCase();
    var recentViews = historyMap[studentToken] || [];
    if (recentViews.length === 0 && lastViewedAt) {
      recentViews = [{ time: lastViewedAt, device: '行動裝置' }];
    }

    students.push({
      className: cls,
      seatNo: seat,
      seatNumber: seat,
      name: name,
      token: token,
      lastViewedAt: lastViewedAt,
      viewCount: viewCount,
      recentViews: recentViews,
      studentId: cls + '-' + seat,
      subjects: [],
      subjectStats: {}
    });
  }

  var sheets = ss.getSheets();
  var subjectList = [];
  var subjectAssessmentsMap = {};

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (!parsed) continue;
    subjectList.push(parsed.sheetName);
    subjectAssessmentsMap[parsed.sheetName] = parsed.assessments || [];

    for (var s = 0; s < students.length; s++) {
      var st = students[s];
      var stSub = parsed.studentMap[st.seatNumber] || parsed.studentMap['NAME_' + st.name];
      if (stSub) {
        st.subjects.push(stSub);
        st.subjectStats[parsed.sheetName] = stSub;
      }
    }
  }

  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  return {
    status: 'success',
    success: true,
    updatedAt: nowStr,
    students: students,
    subjectList: subjectList,
    subjectAssessmentsMap: subjectAssessmentsMap,
    accessLogs: getRecentAccessLogs_(ss, 100)
  };
}

/**
 * 📤 將「學生名冊」的座號與姓名同步至所有科目平時成績分頁（只更新座號與姓名，100% 保留已輸入的各欄分數！）
 */
function syncRosterToSubjectSheets(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet || sSheet.getLastRow() <= 1) {
    if (!silent) ui.alert('【學生名冊】中尚無學生資料！');
    return;
  }

  var sData = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 5).getValues();
  var rosterList = [];
  for (var i = 0; i < sData.length; i++) {
    var seat = normalizeSeatNumber(sData[i][1] || (i + 1));
    var name = (sData[i][2] || '').toString().trim();
    if (seat || name) {
      rosterList.push([seat, name]);
    }
  }

  if (rosterList.length === 0) return;

  var sheets = ss.getSheets();
  var updatedSheets = [];

  for (var idx = 0; idx < sheets.length; idx++) {
    var sh = sheets[idx];
    if (isSystemSheet(sh.getName())) continue;

    var data = ensureAssessmentTypeRow_(sh);
    var loc = locateSubjectSheetRows_(data);
    var startRow = (loc.firstStudentRowIdx !== -1 ? loc.firstStudentRowIdx : 3) + 1; // 1-indexed 學生資料起始列
    sh.getRange(startRow, 1, rosterList.length, 2).setValues(rosterList).setHorizontalAlignment('center');
    updatedSheets.push(sh.getName());
  }

  if (!silent) {
    ui.alert(
      '✅ 姓名同步完成！\n\n' +
      '已將【學生名冊】的 ' + rosterList.length + ' 位學生姓名同步更新至：\n' +
      '【' + updatedSheets.join('、') + '】\n\n' +
      '（各科原有的平時成績分數皆完整保留）'
    );
  }
}

/**
 * 📥 從成績分頁匯入名單至「學生名冊」（含雙向防呆保護與名單預覽確認，絕不誤蓋老師在學生名冊輸入的真實姓名！）
 */
function syncStudentsFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var activeSheet = ss.getActiveSheet();

  // 若老師目前停留在【學生名冊】分頁點擊此功能，代表老師其實是剛改完【學生名冊】，想要把【學生名冊】同步到各科成績表！
  if (activeSheet && activeSheet.getName() === '學生名冊') {
    var resp = ui.alert(
      '💡 請選擇名單同步方向',
      '您目前停留在【學生名冊】分頁。\n\n' +
      '請問您是要將【學生名冊】裡修改好的真實學生姓名，同步更新到「國語平時 / 數學平時 / 社會平時」等各科分頁嗎？\n' +
      '（點選「是」將把學生名冊的姓名寫入各科分頁，且保留原本分數！）',
      ui.ButtonSet.YES_NO
    );
    if (resp === ui.Button.YES) {
      syncRosterToSubjectSheets(false);
    }
    return;
  }

  // 若老師停留在某個科目分頁（例如：國語平時），則以該分頁為來源；否則找第一個科目分頁
  var targetSheet = (!isSystemSheet(activeSheet.getName())) ? activeSheet : null;
  if (!targetSheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (!isSystemSheet(sheets[i].getName())) {
        targetSheet = sheets[i];
        break;
      }
    }
  }

  if (!targetSheet) {
    ui.alert('請先建立至少一個平時成績分頁（例如：國語平時）並填入座號與姓名！');
    return;
  }

  var data = ensureAssessmentTypeRow_(targetSheet);
  var loc = locateSubjectSheetRows_(data);
  var seatCol = loc.seatCol;
  var nameCol = loc.nameCol;
  var dataStartRow = loc.firstStudentRowIdx !== -1 ? loc.firstStudentRowIdx : 2;

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    sSheet = ss.insertSheet('學生名冊', 0);
    sSheet.getRange(1, 1, 1, 5).setValues([['班級', '座號', '姓名', '專屬代碼', '啟用狀態']])
      .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  var existingBySeat = {};
  var fallbackClass = '501';
  var sData = sSheet.getDataRange().getValues();
  for (var s = 1; s < sData.length; s++) {
    var exCls = (sData[s][0] || '').toString().trim();
    var exSeat = normalizeSeatNumber(sData[s][1]);
    var exName = (sData[s][2] || '').toString().trim();
    var exTok = (sData[s][3] || '').toString().trim();
    if (exCls) fallbackClass = exCls;
    if (exSeat) {
      existingBySeat[exSeat] = { className: exCls || fallbackClass, name: exName, token: exTok };
    }
  }

  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  var newRoster = [];

  for (var rowIdx = dataStartRow; rowIdx < data.length; rowIdx++) {
    if (isAssessmentTypeRow_(data[rowIdx], seatCol, nameCol)) continue;
    var seatVal = data[rowIdx][seatCol] !== undefined ? data[rowIdx][seatCol].toString().trim() : '';
    var nameVal = data[rowIdx][nameCol] !== undefined ? data[rowIdx][nameCol].toString().trim() : '';
    if (!seatVal && !nameVal) continue;
    if (seatVal.indexOf('平均') !== -1 || nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal || (newRoster.length + 1));
    var prev = existingBySeat[normSeat] || {};
    var token = prev.token;
    if (!token) {
      token = '';
      for (var t = 0; t < 6; t++) token += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    var clsName = prev.className || fallbackClass;
    newRoster.push([clsName, normSeat, nameVal, token, true]);
  }

  if (newRoster.length === 0) {
    ui.alert('在分頁【' + targetSheet.getName() + '】中未偵測到學生座號與姓名！');
    return;
  }

  // 防呆確認：顯示前 3 位學生姓名，讓老師確認是否真的要用該分頁覆蓋【學生名冊】
  var previewNames = newRoster.slice(0, 3).map(function(r) { return r[1] + '號 ' + r[2]; }).join('、');
  var confirmRes = ui.alert(
    '📥 確認從【' + targetSheet.getName() + '】匯入學生名單？',
    '系統偵測到【' + targetSheet.getName() + '】共有 ' + newRoster.length + ' 位學生：\n' +
    '（前三位：' + previewNames + '...）\n\n' +
    '・點選「是 (YES)」：將用【' + targetSheet.getName() + '】的姓名更新【學生名冊】。\n' +
    '・點選「否 (NO)」：若您是想把【學生名冊】的真實姓名同步去覆蓋各科成績表，請點「否」！',
    ui.ButtonSet.YES_NO
  );

  if (confirmRes !== ui.Button.YES) {
    var reverseRes = ui.alert(
      '📤 是否改為將【學生名冊】同步至各科成績表？',
      '請問您是否要將【學生名冊】目前的學生姓名，反向同步更新到【國語平時 / 數學平時 / 社會平時】等所有科目分頁？',
      ui.ButtonSet.YES_NO
    );
    if (reverseRes === ui.Button.YES) {
      syncRosterToSubjectSheets(false);
    }
    return;
  }

  if (sSheet.getLastRow() > 1) {
    sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 5).clearContent();
  }
  sSheet.getRange(2, 1, newRoster.length, 5).setValues(newRoster);

  sSheet.setColumnWidth(1, 90);
  sSheet.setColumnWidth(2, 80);
  sSheet.setColumnWidth(3, 120);
  sSheet.setColumnWidth(4, 130);
  sSheet.setColumnWidth(5, 100);

  ui.alert('🎉 匯入完成！已從【' + targetSheet.getName() + '】更新 ' + newRoster.length + ' 位學生至【學生名冊】，並完整保留班級名稱與專屬代碼！');
}

/**
 * 🔑 為全班產生「家長查詢代碼」
 */
function generateTokensForStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('學生名冊');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('找不到【學生名冊】工作表！');
    return;
  }
  var data = sheet.getDataRange().getValues();
  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var name = data[i][2];
    var currentToken = (data[i][3] || '').toString().trim();
    if (name && (!currentToken || currentToken.indexOf('DEMO') === 0)) {
      var token = '';
      for (var j = 0; j < 6; j++) {
        token += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      sheet.getRange(i + 1, 4).setValue(token);
      if (data[i][4] === '' || data[i][4] === null) {
        sheet.getRange(i + 1, 5).setValue(true);
      }
      count++;
    }
  }
  SpreadsheetApp.getUi().alert('🔑 已成功為 ' + count + ' 位學生更新專屬 6 碼防偽查詢代碼！');
}

/**
 * ➕ 新增平時成績科目分頁（自動帶入學生名單與預設欄位；教師輸入什麼科目名稱，網站就顯示什麼科目名稱！）
 */
function createNewSubjectSheetFromRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    ui.alert('找不到【學生名冊】分頁，請先執行【🛠️ 建立初始表格】！');
    return;
  }

  var promptResult = ui.prompt(
    '➕ 新增平時成績科目分頁（指定科目名稱）',
    '請輸入科目名稱（例如：國語期中平時、國語期末平時、數學期中平時、自然平時、英文平時）：\n\n💡 您在此輸入什麼科目名稱，家長端網頁與平時成績通知單就會直接顯示該科目名稱！',
    ui.ButtonSet.OK_CANCEL
  );
  if (promptResult.getSelectedButton() !== ui.Button.OK) return;

  var sheetTitle = promptResult.getResponseText().trim();
  if (!sheetTitle) return;
  if (ss.getSheetByName(sheetTitle)) {
    ui.alert('分頁【' + sheetTitle + '】已經存在！');
    return;
  }

  var subPrefix = sheetTitle.charAt(0);
  var foundationPrefix = (sheetTitle.indexOf('數') !== -1) ? '計算' : ((sheetTitle.indexOf('社') !== -1 || sheetTitle.indexOf('自') !== -1) ? '重點' : '語詞');
  var headers = [
    '座號', '姓名',
    subPrefix + '習L1', subPrefix + '習L2', subPrefix + '習L3', subPrefix + '習L4', subPrefix + '習L5', subPrefix + '習L6',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5', '作業卷L6',
    foundationPrefix + 'L1', foundationPrefix + 'L2', foundationPrefix + 'L3', foundationPrefix + 'L4', foundationPrefix + 'L5', foundationPrefix + 'L6',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5', '測驗卷L6'
  ];

  var sData = sSheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < sData.length; i++) {
    var seat = normalizeSeatNumber(sData[i][1]);
    var name = (sData[i][2] || '').toString().trim();
    if (!seat && !name) continue;
    var emptyScores = headers.slice(2).map(function() { return ''; });
    rows.push([seat, name].concat(emptyScores));
  }

  var newSheet = ss.insertSheet(sheetTitle);
  formatDailySubjectSheet(newSheet, sheetTitle + ' 平時成績紀錄', headers, rows);
  ui.alert('🎉 已成功建立科目【' + sheetTitle + '】並自動帶入【學生名冊】全班 ' + rows.length + ' 位學生名單！\n網站與通知小卡將會直接顯示科目名稱：【' + sheetTitle + '】。');
}

/**
 * 📊 查看全班各科平時總均與分項平均統計
 */
function previewClassStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheets = ss.getSheets();
  var summaryLines = [];

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (parsed) {
      summaryLines.push('📘 【' + parsed.sheetName + '】全班平時總平均：' + parsed.classOverallAverage + ' 分');
    }
  }

  if (summaryLines.length === 0) {
    ui.alert('尚未偵測到平時成績分頁！');
  } else {
    ui.alert('📊 全班各科平時成績統計', summaryLines.join('\n\n'), ui.ButtonSet.OK);
  }
}

/**
 * ➕ 新增 1 欄評量項目（方案 C 專用：自動在右側建立新欄、填入今日日期、日曆選取器與作業/考試下拉選單）
 */
function addNewAssessmentColumnMenu() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();
    if (isSystemSheet(sheetName)) {
      SpreadsheetApp.getUi().alert('⚠️ 請先切換至欲新增評量的「科目分頁」（如國語、數學），再執行此功能。');
      return;
    }

    var lastCol = sheet.getLastColumn();
    var targetCol = Math.max(3, lastCol + 1);

    // 格式化新欄位表頭（第 1 列當日日期 + requireDate、第 2 列作業/考試下拉選單、第 3 列項目標題）
    setupSchemeCColumnHeader_(sheet, targetCol, '新評量項目');

    // 選取並聚焦第 3 列標題儲存格，讓教師直接輸入自訂名稱
    sheet.setActiveRange(sheet.getRange(3, targetCol));

    SpreadsheetApp.getActiveSpreadsheet().toast('已成功於第 ' + targetCol + ' 欄建立新評量欄位！已自動填入今日日期與作業/考試下拉選單。', '🎉 新增評量欄位成功', 4);
  } catch (err) {
    SpreadsheetApp.getUi().alert('新增欄位失敗：' + err.toString());
  }
}

/**
 * 針對方案 C 格式化指定欄位的表頭
 * - 第 1 列：自動填入當日日期（或保留自訂日期）＋ 啟用 requireDate 日曆選取器（雙擊彈出日曆）＋ yyyy/MM/dd 格式
 * - 第 2 列：自動建立「作業 / 考試 / 不計分」下拉選單（requireValueInList 顯示下拉箭頭 ▼）＋ 智慧類別判定
 * - 第 3 列：項目標題樣式美化
 * - 最佳化欄寬（84px）與列高
 */
function setupSchemeCColumnHeader_(sheet, colIdx, optItemTitle, optDateVal, optTypeVal) {
  if (!sheet || colIdx < 3) return;
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var now = new Date();
    var nowDateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

    // 1. 取得或設定第 3 列項目標題
    var titleCell = sheet.getRange(3, colIdx);
    var itemTitle = (optItemTitle !== undefined && optItemTitle !== null ? optItemTitle : titleCell.getValue() || '').toString().trim();
    if (!itemTitle) {
      // 若第 3 列為空，檢查第 4 列起是否有任何學生成績
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

    // 若第 1 列、第 2 列、第 3 列皆無任何內容且下方亦無分數，且無傳入參數，則不格式化完全空白的遠端欄位
    var curDateRaw = sheet.getRange(1, colIdx).getValue();
    var curTypeRaw = sheet.getRange(2, colIdx).getValue();
    if (!itemTitle && !curDateRaw && !curTypeRaw && optDateVal === undefined && optTypeVal === undefined && optItemTitle === undefined) {
      return;
    }

    // 2. 設定第 1 列：登記日期（自動帶入今日 + 雙擊彈出日曆選取器）
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

    // 啟用 Google 試算表原生雙擊/點選彈出日曆選取器 (requireDate)
    try {
      var dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(true)
        .build();
      dateCell.setDataValidation(dateRule);
    } catch (dValErr) {}

    // 3. 設定第 2 列：作業 / 考試 下拉選單（requireValueInList + 下拉箭頭 ▼ + 智慧預設類別）
    var typeCell = sheet.getRange(2, colIdx);
    var curTypeVal = (optTypeVal !== undefined && optTypeVal !== null ? optTypeVal : typeCell.getValue() || '').toString().trim();
    if (!curTypeVal || (curTypeVal !== '作業' && curTypeVal !== '考試' && curTypeVal !== '不計分')) {
      curTypeVal = inferDefaultAssessmentType(itemTitle);
      typeCell.setValue(curTypeVal);
    }

    // 啟用在儲存格內顯示下拉箭頭的資料驗證 (showDropdown = true)
    try {
      var typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['作業', '考試', '不計分'], true)
        .setAllowInvalid(true)
        .build();
      typeCell.setDataValidation(typeRule);
    } catch (tValErr) {}

    // 依類別設定底色與字體
    if (curTypeVal === '作業') {
      typeCell.setBackground('#fef3c7').setFontColor('#92400e');
    } else if (curTypeVal === '考試') {
      typeCell.setBackground('#e0f2fe').setFontColor('#075985');
    } else {
      typeCell.setBackground('#f5f5f4').setFontColor('#78716c');
    }
    typeCell.setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    // 4. 設定第 3 列：項目標題樣式美化
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
      titleCell.setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }

    // 5. 設定欄寬與列高
    sheet.setColumnWidth(colIdx, 84);
    sheet.setRowHeight(1, 26);
    sheet.setRowHeight(2, 26);
    sheet.setRowHeight(3, 32);
  } catch (err) {
    Logger.log('setupSchemeCColumnHeader_ 失敗 (col=' + colIdx + '): ' + err.toString());
  }
}

/**
 * 📐 自動整理欄寬與表格排版（內建安全凍結保護，絕不觸發合併儲存格凍結錯誤）
 */
function autoFormatAndResizeAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var count = 0;

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var sName = sh.getName();
    if (sName.startsWith('_')) continue;

    if (sName === '學生名冊') {
      sh.setColumnWidth(1, 90);
      sh.setColumnWidth(2, 80);
      sh.setColumnWidth(3, 120);
      sh.setColumnWidth(4, 130);
      sh.setColumnWidth(5, 100);
      sh.setColumnWidth(6, 150);
      sh.setColumnWidth(7, 110);
      var lr = Math.max(2, sh.getLastRow());
      var maxCol = Math.max(7, sh.getLastColumn());
      sh.getRange(1, 1, lr, maxCol).setVerticalAlignment('middle').setHorizontalAlignment('center');
      safeFreezeSheet_(sh, 1, 0);
      count++;
    } else if (sName === '家長閱覽紀錄') {
      sh.setColumnWidth(1, 150);
      sh.setColumnWidth(2, 80);
      sh.setColumnWidth(3, 80);
      sh.setColumnWidth(4, 120);
      sh.setColumnWidth(5, 120);
      sh.setColumnWidth(6, 140);
      var lrLog = Math.max(2, sh.getLastRow());
      sh.getRange(1, 1, lrLog, 6).setVerticalAlignment('middle').setHorizontalAlignment('center');
      safeFreezeSheet_(sh, 1, 0);
      count++;
    } else if (sName === '成績異動日誌') {
      sh.setColumnWidth(1, 160); // 時間戳記
      sh.setColumnWidth(2, 120); // 科目分頁
      sh.setColumnWidth(3, 140); // 評量項目
      sh.setColumnWidth(4, 110); // 異動類型
      sh.setColumnWidth(5, 140); // 影響學生/數值
      sh.setColumnWidth(6, 200); // 操作來源/備註
      var lrScoreLog = Math.max(2, sh.getLastRow());
      sh.getRange(1, 1, lrScoreLog, 6).setVerticalAlignment('middle').setHorizontalAlignment('center');
      safeFreezeSheet_(sh, 1, 0);
      count++;
    } else {
      ensureAssessmentTypeRow_(sh);
      var maxCol = Math.max(2, sh.getLastColumn());
      var maxRow = Math.max(2, sh.getLastRow());
      sh.setColumnWidth(1, 70);
      sh.setColumnWidth(2, 110);
      for (var c = 3; c <= maxCol; c++) {
        setupSchemeCColumnHeader_(sh, c);
      }
      for (var r = 4; r <= maxRow; r++) {
        sh.setRowHeight(r, 28);
      }
      sh.getRange(1, 1, maxRow, maxCol).setVerticalAlignment('middle').setHorizontalAlignment('center');
      safeFreezeSheet_(sh, 3, 2);
      count++;
    }
  }

  SpreadsheetApp.getUi().alert('🎉 排版完成！已為您將 ' + count + ' 個分頁的欄寬、列高、日曆選取器、下拉選單與凍結窗格最佳化（支援方案 C 版面）！');
}

/**
 * 🔄 將現有成績分頁轉換為「方案 C 版面」
 * 方案 C 規範：
 * - 第 1 列：📌 登記日期（C1.. 填寫日期 YYYY/MM/DD，自動從項目標題/儲存格備註提取或預設今天）
 * - 第 2 列：評量類別（作業 / 考試 / 不計分 下拉選單）
 * - 第 3 列：項目名稱標題列（座號、姓名、國習L1、作業卷L1...）
 * - 第 4 列起：學生成績資料列
 * - 凍結：前 3 列、前 2 欄 (座號、姓名)
 */
function convertAllSubjectSheetsToSchemeC() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '🔄 轉換為「方案 C 版面」',
    '此功能會將所有科目分頁重新整理為「方案 C 版面」：\n\n' +
    '• 第 1 列：📌 登記日期（自動帶入 YYYY/MM/DD）\n' +
    '• 第 2 列：評量類別（作業 / 考試 / 不計分 下拉選單）\n' +
    '• 第 3 列：項目名稱（座號、姓名、評量名稱）\n' +
    '• 第 4 列起：學生成績資料\n' +
    '• 凍結前 3 列與前 2 欄\n\n' +
    '確定要轉換所有科目成績分頁嗎？',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var convertedCount = 0;
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
  var todayDateStr = Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sName = sheet.getName();
    if (isSystemSheet(sName)) continue;

    try {
      var data = sheet.getDataRange().getValues();
      if (!data || data.length < 2) continue;

      var loc = locateSubjectSheetRows_(data);
      if (loc.itemHeaderRowIdx === -1) continue;

      var seatCol = loc.seatCol;
      var nameCol = loc.nameCol;
      var itemHeaderRow = data[loc.itemHeaderRowIdx];
      var typeRow = (loc.typeRowIdx !== -1 && loc.typeRowIdx < data.length) ? data[loc.typeRowIdx] : null;
      var dateRow = (loc.dateRowIdx !== -1 && loc.dateRowIdx < data.length) ? data[loc.dateRowIdx] : null;

      var headers = ['座號', '姓名'];
      var dateValues = [];
      var typeValues = ['評量', '作業/考試'];
      var colIndices = [];

      for (var c = 0; c < itemHeaderRow.length; c++) {
        if (c === seatCol || c === nameCol) continue;
        var hName = (itemHeaderRow[c] || '').toString().trim();
        if (!hName) continue;

        headers.push(hName);
        colIndices.push(c);

        // 提取日期：優先 dateRow，次之 note，再次之標題，最後為今天
        var colDateVal = '';
        if (dateRow && dateRow[c]) {
          var pD = extractDateFromTitle_(dateRow[c].toString());
          if (pD) colDateVal = Utilities.formatDate(pD, tz, 'yyyy/MM/dd');
        }
        if (!colDateVal) {
          try {
            var note = sheet.getRange(loc.itemHeaderRowIdx + 1, c + 1).getNote();
            if (note) {
              var nD = extractDateFromTitle_(note);
              if (nD) colDateVal = Utilities.formatDate(nD, tz, 'yyyy/MM/dd');
            }
          } catch (eNote) {}
        }
        if (!colDateVal) {
          var tD = extractDateFromTitle_(hName);
          if (tD) colDateVal = Utilities.formatDate(tD, tz, 'yyyy/MM/dd');
        }
        if (!colDateVal) {
          colDateVal = todayDateStr;
        }
        dateValues.push(colDateVal);

        // 提取評量類別
        var aType = '';
        if (typeRow && typeRow[c]) {
          aType = (typeRow[c] || '').toString().trim();
        }
        if (!aType || (aType !== '作業' && aType !== '考試' && aType !== '不計分')) {
          aType = inferDefaultAssessmentType(hName);
        }
        typeValues.push(aType);
      }

      // 提取學生成績列
      var studentRows = [];
      var firstStudentIdx = loc.firstStudentRowIdx !== -1 ? loc.firstStudentRowIdx : 3;
      for (var r = firstStudentIdx; r < data.length; r++) {
        var rSeat = (data[r][seatCol] || '').toString().trim();
        var rName = (data[r][nameCol] || '').toString().trim();
        if (!rSeat && !rName) continue;

        var rowArr = [rSeat, rName];
        for (var k = 0; k < colIndices.length; k++) {
          var scoreVal = data[r][colIndices[k]];
          rowArr.push(scoreVal !== undefined ? scoreVal : '');
        }
        studentRows.push(rowArr);
      }

      formatDailySubjectSheet(sheet, sName, headers, studentRows);

      // 回填原先的自訂日期與類別
      if (dateValues.length > 0) {
        var dRange = sheet.getRange(1, 3, 1, dateValues.length);
        dRange.setValues([dateValues]).setNumberFormat('yyyy/MM/dd');
        try {
          var dRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
          dRange.setDataValidation(dRule);
        } catch (eDateRule) {}
      }
      if (typeValues.length > 2) {
        var tRange = sheet.getRange(2, 1, 1, typeValues.length);
        tRange.setValues([typeValues]);
        try {
          var tRule = SpreadsheetApp.newDataValidation().requireValueInList(['作業', '考試', '不計分'], true).setAllowInvalid(true).build();
          sheet.getRange(2, 3, 1, typeValues.length - 2).setDataValidation(tRule);
        } catch (eTypeRule) {}
      }

      logScoreUpdateRecord_(sName, '全部分頁評量', '轉換為方案 C 版面', studentRows.length + ' 位學生', '選單一鍵轉換');
      convertedCount++;
    } catch (err) {
      Logger.log('轉換科目分頁 ' + sName + ' 失敗: ' + err.toString());
    }
  }

  ui.alert('完成', '🎉 已成功將 ' + convertedCount + ' 個科目分頁轉換為「方案 C 版面」！\n\n• 第 1 列：📌 登記日期\n• 第 2 列：作業 / 考試 選項\n• 第 3 列：項目名稱\n• 第 4 列起：學生成績', ui.ButtonSet.OK);
}

/**
 * 接收網站快貼小幫手發送之平時成績並寫入分頁
 */
function appendDailyScores(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var payload = typeof body === 'string' ? JSON.parse(body) : body;
  var sheetName = (payload.sheetName || '國語平時').trim();
  var headers = payload.headers || ['座號', '姓名', '習作L1', '作業卷L1', '測驗卷L1'];
  var rows = payload.rows || [];

  if (rows.length === 0) {
    return { status: 'error', success: false, message: '沒有平時成績資料可儲存' };
  }

  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  formatDailySubjectSheet(sheet, sheetName + ' 平時成績紀錄', headers, rows);

  return {
    status: 'success',
    success: true,
    message: '成功建立並更新【' + sheetName + '】，共 ' + rows.length + ' 位學生平時成績！',
    sheetName: sheetName,
    studentCount: rows.length
  };
}

// 相容別名（確保無論舊版或新版 Loader.gs 呼叫皆 100% 正常運作）
function handleGetRequest(e) { return doGet(e); }
function handlePostRequest(e) { return doPost(e); }
function generateOrUpdateStudentTokens() { return generateTokensForStudents(); }
function previewClassScoreSummary() { return previewClassStatistics(); }
function createDemoDailySheets() { return setupStandardSheets(); }

// =========================================================================
// =========================================================================
// 8. School LINE Hub 中樞接入模組（免 Token、免跨表、統一推播與綁定管理）
// =========================================================================

/**
 * School LINE Hub 客戶端 SDK (若環境中尚未定義則初始化)
 */
var LineHubClient = this.LineHubClient || (function() {
  function Client(config) {
    if (!config || !config.hubUrl || !config.appId || !config.appSecret) {
      throw new Error("LineHubClient 初始化失敗：必須提供 hubUrl, appId 與 appSecret！");
    }
    this.hubUrl = config.hubUrl.trim();
    this.appId = config.appId.trim();
    this.appSecret = config.appSecret.trim();
    this.timeout = config.timeout || 30000;
  }

  Client.prototype._post = function(action, extraPayload) {
    var payload = Object.assign({
      action: action,
      appId: this.appId,
      appSecret: this.appSecret
    }, extraPayload || {});

    var jsonStr = JSON.stringify(payload);

    if (typeof UrlFetchApp !== 'undefined') {
      try {
        var resp = UrlFetchApp.fetch(this.hubUrl, {
          method: 'post',
          contentType: 'application/json',
          payload: jsonStr,
          muteHttpExceptions: true
        });
        var code = resp.getResponseCode();
        var body = resp.getContentText();
        var parsed = {};
        try {
          parsed = JSON.parse(body);
        } catch (e) {
          parsed = { status: 'error', code: code, raw: body };
        }
        return parsed;
      } catch (gasErr) {
        return { status: 'error', message: gasErr.toString() };
      }
    }

    if (typeof fetch !== 'undefined') {
      return fetch(this.hubUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: jsonStr
      }).then(function(res) {
        return res.json();
      }).catch(function(err) {
        return { status: 'error', message: err.toString() };
      });
    }

    throw new Error("找不到可用的 HTTP 客戶端 (UrlFetchApp 或 fetch)！");
  };

  Client.prototype.testConnection = function() {
    return this._post('verify_client');
  };

  Client.prototype.getStudentBinding = function(target) {
    return this._post('get_student_binding', { target: target });
  };

  Client.prototype.getAllBindingsSummary = function() {
    return this._post('get_all_bindings_summary');
  };

  Client.prototype.registerStudents = function(studentsList) {
    return this._post('register_students', { students: studentsList });
  };

  Client.prototype.sendPushNotification = function(options) {
    return this._post('send_push', {
      target: {
        seatNo: options.seatNo,
        studentId: options.studentId,
        studentName: options.studentName,
        token: options.token,
        title: options.title,
        body: options.body,
        url: options.url,
        customMessage: options.customMessage
      }
    });
  };

  Client.prototype.sendBatchPush = function(targets, messageTemplate) {
    return this._post('send_push', {
      targets: targets,
      messageTemplate: messageTemplate
    });
  };

  return Client;
})();

/**
 * 預設 School LINE Hub 連線組態
 */
var SCHOOL_LINE_HUB_DEFAULT_CONFIG = {
  hubUrl: 'https://script.google.com/macros/s/AKfycby4yrSdmztcbnmdHlGdLqqM9nBRT4vo6DwK8LmW1R-wi22qEKC_vcD-VvrDnUTp4Bm_Yg/exec',
  appId: 'PROJECT_SCORE',
  appSecret: 'sec_score_demo_7182934'
};

/**
 * 取得目前 School LINE Hub 連線組態（支援屬性覆寫，預設即為官方中樞）
 */
function getLineHubConfig_() {
  var prop = PropertiesService.getScriptProperties();
  var hubUrl = (prop.getProperty('LINE_HUB_URL') || '').trim() || SCHOOL_LINE_HUB_DEFAULT_CONFIG.hubUrl;
  var appId = (prop.getProperty('LINE_HUB_APP_ID') || '').trim() || SCHOOL_LINE_HUB_DEFAULT_CONFIG.appId;
  var appSecret = (prop.getProperty('LINE_HUB_APP_SECRET') || '').trim() || SCHOOL_LINE_HUB_DEFAULT_CONFIG.appSecret;

  return {
    hubUrl: hubUrl,
    appId: appId,
    appSecret: appSecret
  };
}

/**
 * 實例化 LineHubClient 客戶端
 */
function getLineHubClient_() {
  return new LineHubClient(getLineHubConfig_());
}

/**
 * 於【學生名冊】第 1 列建立 School LINE Hub 相關標籤儲存格
 */
function initLineSettingsCellsOnRoster_(sSheet) {
  if (!sSheet) return;
  try {
    var l1Val = (sSheet.getRange(1, 12).getValue() || '').toString().trim();
    if (l1Val !== '【School LINE Hub】') {
      sSheet.getRange(1, 12).setValue('【School LINE Hub】')
        .setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#047857').setHorizontalAlignment('center');
      sSheet.getRange(1, 13).setValue('PROJECT_SCORE (中樞管理)')
        .setFontFamily('monospace').setFontSize(9).setHorizontalAlignment('center');
      sSheet.setColumnWidth(11, 25);
      sSheet.setColumnWidth(12, 140);
      sSheet.setColumnWidth(13, 160);
    }

    var n1Val = (sSheet.getRange(1, 14).getValue() || '').toString().trim();
    if (n1Val !== '【中樞連線狀態】') {
      sSheet.getRange(1, 14).setValue('【中樞連線狀態】')
        .setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#0369a1').setHorizontalAlignment('center');
      sSheet.getRange(1, 15).setValue('ACTIVE')
        .setFontFamily('monospace').setFontSize(9).setHorizontalAlignment('center');
      sSheet.setColumnWidth(14, 130);
      sSheet.setColumnWidth(15, 120);
    }
  } catch (err) {}
}

/**
 * 遮罩學生姓名（姓名第 2 字改為 ○）保護個資
 */
function maskStudentName_(name) {
  if (!name) return "";
  var str = name.toString().trim();
  if (str.length === 2) return str.substring(0, 1) + "○";
  if (str.length >= 3) return str.substring(0, 1) + "○" + str.substring(2);
  return str;
}

/**
 * 自評量欄位標題解析日期（支援 YYYY-MM-DD, M/D, M-D, M月D日 等格式）
 */
function extractDateFromTitle_(title, referenceYear) {
  if (!title) return null;
  if (title instanceof Date) return title;
  var str = title.toString().trim();
  if (str.indexOf('GMT') !== -1 || str.indexOf('台北標準時間') !== -1 || str.indexOf('UTC') !== -1) {
    var parsed = new Date(str);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  referenceYear = referenceYear || (new Date()).getFullYear();

  // Pattern 1: YYYY-MM-DD or YYYY/MM/DD
  var m1 = str.match(/(20\d{2})[\/\-\.](0?[1-9]|1[0-2])[\/\-\.](0?[1-9]|[12]\d|3[01])/);
  if (m1) {
    return new Date(parseInt(m1[1], 10), parseInt(m1[2], 10) - 1, parseInt(m1[3], 10));
  }

  // Pattern 2: MM/DD or MM-DD or M/D (例如 10/24, 10-25)
  var m2 = str.match(/(0?[1-9]|1[0-2])[\/\-\.](0?[1-9]|[12]\d|3[01])/);
  if (m2) {
    return new Date(referenceYear, parseInt(m2[1], 10) - 1, parseInt(m2[2], 10));
  }

  // Pattern 3: M月D日 或 MM月DD日
  var m3 = str.match(/(0?[1-9]|1[0-2])\s*月\s*(0?[1-9]|[12]\d|3[01])\s*日?/);
  if (m3) {
    return new Date(referenceYear, parseInt(m3[1], 10) - 1, parseInt(m3[2], 10));
  }

  return null;
}

/**
 * 依學生代碼與座號，自各科目成績分頁提取即時成績摘要 (Score Summary)
 * 支援方案 B：依自訂項目計數 (subjectCounts) 或基準日期 (sinceDate) 計算各科更新項目數
 * 送出給 School LINE Hub 作為推播內文
 */
function generateStudentScoreSummary_(studentToken, studentSeat, studentName, options) {
  options = options || {};
  var summaryLines = [];
  var customCounts = options.subjectCounts || null;
  if (typeof customCounts === 'string') {
    try {
      customCounts = JSON.parse(customCounts);
    } catch (e) {
      customCounts = null;
    }
  }

  // 1. 若有指定各科更新項目計數 (subjectCounts)
  if (customCounts && typeof customCounts === 'object') {
    for (var subj in customCounts) {
      if (customCounts.hasOwnProperty(subj)) {
        var cnt = parseInt(customCounts[subj], 10);
        if (!isNaN(cnt) && cnt > 0) {
          summaryLines.push('• ' + subj + '：有（ ' + cnt + ' ）項成績更新');
        }
      }
    }
  }

  // 2. 若無指定各科項目計數，則自動自試算表各科目分頁提取（並支援 sinceDate 日期過濾）
  if (summaryLines.length === 0) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheets = ss.getSheets();
      var sinceDateStr = (options.sinceDate || '').toString().trim().replace(/\//g, '-');

      for (var i = 0; i < sheets.length; i++) {
        var sh = sheets[i];
        if (isSystemSheet(sh.getName())) continue;
        var parsed = parseDailySubjectSheet(sh);
        if (parsed) {
          var sSeat = normalizeSeatNumber(studentSeat);
          var stSubj = parsed.studentMap[sSeat] || parsed.studentMap['NAME_' + studentName];
          var count = 0;

          var colsToScan = (parsed.itemCols && parsed.itemCols.length > 0) ? parsed.itemCols : (parsed.assessments || []);

          if (sinceDateStr && colsToScan.length > 0) {
            for (var a = 0; a < colsToScan.length; a++) {
              var aItem = colsToScan[a];
              var aDate = extractDateFromTitle_(aItem.date) || extractDateFromTitle_(aItem.title || aItem.name);
              if (aDate) {
                var aDateStr = Utilities.formatDate(aDate, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd');
                if (aDateStr >= sinceDateStr) {
                  count++;
                }
              }
            }
          }

          if (count === 0 && !sinceDateStr) {
            if (stSubj && stSubj.totalScoresCount > 0) {
              count = stSubj.totalScoresCount;
            } else {
              count = 1;
            }
          }

          if (count > 0 || !sinceDateStr) {
            var subjName = parsed.subjectShortName || parsed.sheetName || sh.getName();
            summaryLines.push('• ' + subjName + '：有（ ' + (count > 0 ? count : 1) + ' ）項成績更新');
          }
        }
      }
    } catch (e) {
      Logger.log('計算成績摘要失敗: ' + e.toString());
    }
  }

  if (summaryLines.length === 0) {
    return '目前平時評量成績摘要：\n• 平時成績：有（ 1 ）項成績更新';
  }

  return '目前平時評量成績摘要：\n' + summaryLines.join('\n');
}

/**
 * 試算表即時編輯觸發器（方案 B & 方案 C 核心連動）：
 * 1. 當在第 3 列新增/修改評量名稱時（例如輸入 L6測驗卷）：
 *    - 第 1 列自動補齊當天日期（YYYY/MM/DD），並啟用 double-click 彈出日曆選取器 (requireDate)
 *    - 第 2 列自動產生「作業 / 考試 / 不計分」下拉選單（含下拉箭頭 ▼），並智慧帶入預設類別與顏色
 *    - 第 3 列項目標題自動美化底色，欄寬自動最佳化 (84px)
 * 2. 當編輯第 4 列起之學生成績時：
 *    - 自動確保該欄第 1 列登記日期、第 2 列下拉選單與第 3 列備註齊全
 *    - 自動寫入【成績異動日誌】分頁（時間、科目、項目、類型、學生/數值、來源）
 *    - 同步記錄至 DocumentProperties，供推播動態抓取試算表更新時間
 * 3. 當點選日曆或手動修改登記日期（第 1 列）或評量類別（第 2 列）時，亦記錄至異動日誌
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    if (isSystemSheet(sheetName)) return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (col < 3) return;

    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    var now = new Date();
    var nowStr = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    var nowDateStr = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

    // 取得項目名稱（第 3 列）
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
      // 確保該欄的表頭具備日期、日曆選取器與作業/考試下拉選單
      setupSchemeCColumnHeader_(sheet, col, itemTitle);

      // 於項目標題儲存格（第 3 列）加入或更新備註
      try {
        var headerCell = sheet.getRange(3, col);
        var oldNote = headerCell.getNote() || '';
        var cleanNote = oldNote.replace(/最後更新：[^\n]+/g, '').trim();
        headerCell.setNote((cleanNote ? cleanNote + '\n' : '') + '最後更新：' + nowStr);
      } catch (notErr) {}

      // 記錄至 DocumentProperties
      try {
        var docProps = PropertiesService.getDocumentProperties();
        docProps.setProperty('EDIT_' + sheetName + '_' + col, nowStr);
      } catch (prpErr) {}

      // 方案 B：自動寫入【成績異動日誌】
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
  } catch (err) {
    Logger.log('onEdit 執行失敗: ' + err.toString());
  }
}

/**
 * 取得 LINE 推播與 School LINE Hub 中樞綁定狀態摘要（供網頁教師端與選單查詢）
 * 支援方案 B：接收 sinceDate 參數，主動抓取試算表各科評量與儲存格更新時間，動態計算更新項目數
 */
function getLinePushStatus_(params) {
  params = params || {};
  var hubClient = getLineHubClient_();
  var config = getLineHubConfig_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet || sSheet.getLastRow() <= 1) {
    return {
      status: 'success',
      success: true,
      hasToken: true,
      hasSourceUrl: true,
      hasHub: true,
      hubConnected: true,
      hubUrl: config.hubUrl,
      appId: config.appId,
      totalStudents: 0,
      boundStudentsCount: 0,
      boundParentsCount: 0,
      students: []
    };
  }

  // 1. 測試中樞連線
  var testRes = {};
  try {
    testRes = hubClient.testConnection();
  } catch (err) {
    testRes = { status: 'error', message: err.toString() };
  }

  // 2. 批次取得全班綁定對照表
  var summaryRes = {};
  var bindingsMap = {};
  try {
    summaryRes = hubClient.getAllBindingsSummary();
    if (summaryRes && summaryRes.bindings) {
      bindingsMap = summaryRes.bindings;
    }
  } catch (err) {
    Logger.log('取得中樞綁定摘要失敗: ' + err.toString());
  }

  var rData = sSheet.getDataRange().getValues();
  var studentsList = [];
  var totalBoundParents = 0;
  var boundStudentsCount = 0;

  for (var i = 1; i < rData.length; i++) {
    var rawSeat = (rData[i][1] || '').toString().trim();
    var seat = normalizeSeatNumber(rawSeat);
    var name = (rData[i][2] || '').toString().trim();
    var tok = (rData[i][3] || '').toString().trim().toUpperCase();
    var isEnabled = rData[i][4] !== false;

    if (!name || !tok || !isEnabled) continue;

    // 交叉比對中樞綁定：支援依座號（01 或 1）、查詢代碼 (Token)
    var strippedSeat = rawSeat.replace(/^0+/, '');
    var bInfo = bindingsMap[seat] || bindingsMap[strippedSeat] || bindingsMap[tok] || null;
    var isBound = bInfo ? (bInfo.isBound !== false) : false;
    var pCount = isBound ? 1 : 0;

    if (pCount > 0) {
      boundStudentsCount++;
      totalBoundParents += pCount;
    }

    studentsList.push({
      seat: seat,
      name: name,
      token: tok,
      boundParentsCount: pCount
    });
  }

  var hubConnected = testRes && (testRes.status === 'success' || summaryRes.status === 'success');

  // 3. 取得前次推播紀錄與各科目資訊（方案 B 基準追蹤）
  var props = PropertiesService.getScriptProperties();
  var lastPushTimestamp = props.getProperty('LAST_SCORE_PUSH_TIMESTAMP') || '';
  var lastPushDate = props.getProperty('LAST_SCORE_PUSH_DATE') || '';
  if (!lastPushDate && lastPushTimestamp) {
    lastPushDate = lastPushTimestamp.split(' ')[0];
  }

  // 採計基準日期：優先使用前端傳入之設定日期 sinceDate，若無則退回使用前次推播日期 lastPushDate
  var rawSinceDate = (params.sinceDate || '').toString().trim() || lastPushDate;
  var targetSinceDate = rawSinceDate.replace(/\//g, '-');

  // 抓取試算表檔案最後修改時間 (Last Updated)
  var ssLastUpdatedStr = '';
  try {
    var file = DriveApp.getFileById(ss.getId());
    var fDate = file.getLastUpdated();
    ssLastUpdatedStr = Utilities.formatDate(fDate, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  } catch (dErr) {}

  var docProps = null;
  try {
    docProps = PropertiesService.getDocumentProperties();
  } catch (e) {}

  // 方案 B：預先讀取【成績異動日誌】分頁之最新紀錄
  var scoreLogMap = {};
  try {
    var logSheet = ss.getSheetByName('成績異動日誌');
    if (logSheet && logSheet.getLastRow() > 1) {
      var logData = logSheet.getRange(2, 1, Math.min(logSheet.getLastRow() - 1, 300), 3).getValues();
      for (var lIdx = 0; lIdx < logData.length; lIdx++) {
        var logTimeStr = logData[lIdx][0];
        var logSubj = (logData[lIdx][1] || '').toString().trim();
        var logItem = (logData[lIdx][2] || '').toString().trim();
        if (logTimeStr && logSubj && logItem) {
          var logD = extractDateFromTitle_(logTimeStr);
          if (logD) {
            var lKey = logSubj + '_' + logItem;
            if (!scoreLogMap[lKey] || logD > scoreLogMap[lKey]) {
              scoreLogMap[lKey] = logD;
            }
          }
        }
      }
    }
  } catch (logMapErr) {}

  var subjectsMeta = [];
  try {
    var allSheets = ss.getSheets();
    for (var sIdx = 0; sIdx < allSheets.length; sIdx++) {
      var sh = allSheets[sIdx];
      if (isSystemSheet(sh.getName())) continue;
      var parsed = parseDailySubjectSheet(sh);
      if (parsed) {
        var colsToScan = (parsed.itemCols && parsed.itemCols.length > 0) ? parsed.itemCols : (parsed.assessments || []);
        var updatedCount = 0;
        var latestColDateStr = '';

        // 遍歷所有評量欄位，抓取試算表的更新時間
        for (var aIdx = 0; aIdx < colsToScan.length; aIdx++) {
          var item = colsToScan[aIdx];
          var colIdx1 = item.colIdx !== undefined ? (item.colIdx + 1) : (aIdx + 3);
          var colDate = null;
          var colDateStr = '';

          // 0. 方案 C：優先自第 1 列登記日期 (item.date) 提取
          if (item.date) {
            var directDate = extractDateFromTitle_(item.date);
            if (directDate) {
              colDate = directDate;
            }
          }

          // 0.1 若物件中無日期，直接讀取試算表第 1 列儲存格
          if (!colDate) {
            try {
              var r1Val = sh.getRange(1, colIdx1).getValue();
              if (r1Val) {
                var r1Date = extractDateFromTitle_(r1Val);
                if (r1Date) colDate = r1Date;
              }
            } catch (r1Err) {}
          }

          // 1. 嘗試從【成績異動日誌】抓取更新時間（方案 B 核心）
          var itemTitleName = item.title || item.name || '';
          if (!colDate && scoreLogMap && itemTitleName) {
            var sKey = sh.getName() + '_' + itemTitleName;
            if (scoreLogMap[sKey]) {
              colDate = scoreLogMap[sKey];
            }
          }

          // 2. 嘗試從欄位標題解析日期 (例如 10/24, 2026-10-28, 10月29日)
          if (!colDate && itemTitleName) {
            var titleDate = extractDateFromTitle_(itemTitleName);
            if (titleDate) {
              colDate = titleDate;
            }
          }

          // 3. 嘗試從儲存格備註 (Note) 抓取更新時間（檢查第 3 列或第 2 列標題儲存格）
          if (!colDate) {
            try {
              var headerCell = sh.getRange(3, colIdx1);
              var note = headerCell.getNote();
              if (!note) {
                headerCell = sh.getRange(2, colIdx1);
                note = headerCell.getNote();
              }
              if (note) {
                var noteDate = extractDateFromTitle_(note);
                if (noteDate) colDate = noteDate;
              }
            } catch (nErr) {}
          }

          // 4. 嘗試從 DocumentProperties 抓取記錄的欄位編輯時間
          if (!colDate && docProps) {
            try {
              var propTime = docProps.getProperty('EDIT_' + sh.getName() + '_' + colIdx1);
              if (propTime) {
                var propDate = extractDateFromTitle_(propTime);
                if (propDate) colDate = propDate;
              }
            } catch (pErr) {}
          }

          if (colDate) {
            colDateStr = Utilities.formatDate(colDate, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd');
            if (!latestColDateStr || colDateStr > latestColDateStr) {
              latestColDateStr = colDateStr;
            }
          }

          // 依設定日期 (targetSinceDate) 判定是否屬於更新項目
          if (targetSinceDate) {
            if (colDateStr && colDateStr >= targetSinceDate) {
              updatedCount++;
            }
          } else {
            updatedCount++;
          }
        }

        // 若該科目有評量，但欄位標題無日期：檢驗試算表檔案最後修改時間 (ssLastUpdatedStr)
        if (updatedCount === 0 && colsToScan.length > 0) {
          if (ssLastUpdatedStr && targetSinceDate) {
            var fileDateStr = ssLastUpdatedStr.split(' ')[0];
            if (fileDateStr >= targetSinceDate) {
              updatedCount = 1;
              latestColDateStr = fileDateStr;
            }
          } else if (!targetSinceDate) {
            updatedCount = 1;
          }
        }

        var cleanSubjName = parsed.subjectShortName || parsed.sheetName || sh.getName();
        subjectsMeta.push({
          name: cleanSubjName,
          sheetName: sh.getName(),
          totalAssessments: colsToScan.length,
          defaultCount: updatedCount,
          latestUpdate: latestColDateStr || (ssLastUpdatedStr ? ssLastUpdatedStr.split(' ')[0] : '')
        });
      }
    }
  } catch (metaErr) {
    Logger.log('取得科目評量資訊失敗: ' + metaErr.toString());
  }

  return {
    status: 'success',
    success: true,
    hasToken: hubConnected,
    hasSourceUrl: hubConnected,
    hasHub: true,
    hubConnected: hubConnected,
    hubUrl: config.hubUrl,
    appId: config.appId,
    projectName: testRes.projectName || '平時成績管理系統',
    totalStudents: rData.length - 1,
    boundStudentsCount: boundStudentsCount,
    boundParentsCount: totalBoundParents,
    students: studentsList,
    sinceDate: targetSinceDate,
    spreadsheetLastUpdated: ssLastUpdatedStr,
    lastPushTimestamp: lastPushTimestamp,
    lastPushDate: lastPushDate,
    subjectsMeta: subjectsMeta,
    diagnostic: {
      hubUrl: config.hubUrl,
      appId: config.appId,
      hubConnected: hubConnected,
      hubMessage: testRes.message || (hubConnected ? '已成功連線 School LINE Hub 中樞' : (testRes.message || '中樞連線逾時')),
      activeBindingsCount: testRes.activeBindingsCount || totalBoundParents
    }
  };
}

/**
 * 執行平時成績更新推播（接入 School LINE Hub 中樞）
 * 支援方案 B：接收 sinceDate 與 subjectCounts，動態組裝符合規範之成績通知訊息
 */
function executeLineScorePushNotifications_(params) {
  params = params || {};
  var hubClient = getLineHubClient_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet || sSheet.getLastRow() <= 1) {
    return {
      status: 'error',
      success: false,
      code: 'NO_ROSTER',
      message: '查無學生名冊資料！'
    };
  }

  var rData = sSheet.getDataRange().getValues();
  var allStudents = [];

  for (var i = 1; i < rData.length; i++) {
    var rawSeat = (rData[i][1] || '').toString().trim();
    var seat = normalizeSeatNumber(rawSeat);
    var name = (rData[i][2] || '').toString().trim();
    var tok = (rData[i][3] || '').toString().trim().toUpperCase();
    var isEnabled = rData[i][4] !== false;

    if (!name || !tok || !isEnabled) continue;

    allStudents.push({
      seat: seat,
      rawSeat: rawSeat,
      name: name,
      token: tok
    });
  }

  // 支援個別指定學生發送（傳入 token 或 tokens 逗號分隔列表）
  var filterTokens = [];
  if (params.token) {
    filterTokens.push(params.token.toString().trim().toUpperCase());
  }
  if (params.tokens) {
    var splitToks = params.tokens.toString().split(',');
    for (var tk = 0; tk < splitToks.length; tk++) {
      var itemTok = splitToks[tk].trim().toUpperCase();
      if (itemTok && filterTokens.indexOf(itemTok) === -1) {
        filterTokens.push(itemTok);
      }
    }
  }

  var studentsToNotify = allStudents;
  if (filterTokens.length > 0) {
    studentsToNotify = allStudents.filter(function(st) {
      return filterTokens.indexOf(st.token) !== -1;
    });
  }

  if (studentsToNotify.length === 0) {
    return {
      status: 'error',
      success: false,
      code: 'NO_STUDENTS',
      message: '查無符合條件的學生資料！'
    };
  }

  var gasUrl = (params.gasUrl || params.api || '').toString().trim() || ScriptApp.getService().getUrl() || '';
  var parentBaseUrl = (params.parentBaseUrl || '').toString().trim() || 'https://chy23.github.io/MPScore-Daily/parent.html';
  if (parentBaseUrl && !parentBaseUrl.endsWith('.html')) {
    parentBaseUrl = (parentBaseUrl.endsWith('/') ? parentBaseUrl : parentBaseUrl + '/') + 'parent.html';
  }

  var summaryOptions = {
    sinceDate: params.sinceDate || '',
    subjectCounts: params.subjectCounts || null
  };

  // 組裝推播目標，直接送出座號、姓名與成績摘要給中樞發送
  var targets = [];
  for (var stIdx = 0; stIdx < studentsToNotify.length; stIdx++) {
    var stItem = studentsToNotify[stIdx];
    var studentUrl = parentBaseUrl + '?t=' + encodeURIComponent(stItem.token);
    if (gasUrl) {
      studentUrl += '&api=' + encodeURIComponent(gasUrl);
    }

    // 計算各科成績摘要
    var scoreSummary = generateStudentScoreSummary_(stItem.token, stItem.seat, stItem.name, summaryOptions);
    var maskedName = maskStudentName_(stItem.name);

    var customMsg = '📊 【平時成績更新通知】\n' +
      maskedName + ' 家長您好：\n' +
      scoreSummary + '\n' +
      '🔗 點此開啟明細查看：\n' +
      studentUrl + '\n' +
      '（由 平時成績管理系統 自動發送）';

    targets.push({
      seatNo: stItem.seat,
      studentName: stItem.name,
      token: stItem.token,
      title: '📊 【平時成績更新通知】',
      body: scoreSummary,
      url: studentUrl,
      customMessage: customMsg
    });
  }

  var template = {
    title: '📊 【平時成績更新通知】',
    body: '平時評量成績已完成登錄，請家長參閱最新學習手札。'
  };

  // 呼叫 LineHubClient 發送批次推播
  var hubResult = {};
  try {
    hubResult = hubClient.sendBatchPush(targets, template);
  } catch (sendErr) {
    return {
      status: 'error',
      success: false,
      message: 'School LINE Hub 中樞發送失敗: ' + sendErr.toString()
    };
  }

  var sentCount = (hubResult && typeof hubResult.sentCount === 'number') ? hubResult.sentCount : 0;
  var failedCount = (hubResult && typeof hubResult.failedCount === 'number') ? hubResult.failedCount : 0;
  var skippedCount = (hubResult && typeof hubResult.skippedCount === 'number') ? hubResult.skippedCount : 0;

  var isSuccess = (sentCount > 0);
  var nowFormatted = '';
  var nowDateFormatted = '';

  // 若發送成功，記錄本次推播時間戳記作為下一次的「前次推播時間」
  if (isSuccess) {
    var now = new Date();
    var tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    nowFormatted = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss');
    nowDateFormatted = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    try {
      var saveProps = PropertiesService.getScriptProperties();
      saveProps.setProperty('LAST_SCORE_PUSH_TIMESTAMP', nowFormatted);
      saveProps.setProperty('LAST_SCORE_PUSH_DATE', nowDateFormatted);
    } catch (saveErr) {
      Logger.log('儲存推播時間失敗: ' + saveErr.toString());
    }
  }

  var msg = '';
  if (isSuccess) {
    msg = '成功透過 School LINE Hub 發送 ' + sentCount + ' 則成績通知！';
  } else if (hubResult && hubResult.message) {
    msg = hubResult.message;
  } else if (skippedCount > 0) {
    msg = '未發送：所選學生尚未綁定家長 LINE（請先於 LINE 完成綁定）';
  } else {
    msg = '發送失敗：中樞未送出任何訊息，請檢查連線狀態。';
  }

  return {
    status: isSuccess ? 'success' : 'error',
    success: isSuccess,
    successCount: sentCount,
    failCount: failedCount,
    skippedCount: skippedCount,
    studentCount: studentsToNotify.length,
    lastPushTimestamp: nowFormatted,
    lastPushDate: nowDateFormatted,
    hubDetails: hubResult.details || [],
    message: msg
  };
}

/**
 * ⚙️ 設定 School LINE Hub 中樞連線（選單對話框）
 */
function setupLinePushSettingsMenu() {
  var ui = SpreadsheetApp.getUi();
  var config = getLineHubConfig_();

  var hubClient = getLineHubClient_();
  var testRes = {};
  try {
    testRes = hubClient.testConnection();
  } catch (e) {
    testRes = { status: 'error', message: e.toString() };
  }

  var isOk = testRes && testRes.status === 'success';

  var statusMsg = '【School LINE Hub 中樞接入資訊】\n\n' +
    '• 中樞網址：' + config.hubUrl + '\n' +
    '• 專案 App ID：' + config.appId + '\n' +
    '• 連線狀態：' + (isOk ? '🟢 連線正常（' + (testRes.projectName || '平時成績管理系統') + '）' : '⚠️ 連線異常 (' + (testRes.message || '請確認中樞已部署') + ')') + '\n\n' +
    '本系統已全面接入 School LINE Hub 統一中樞，無需再設定 LINE Channel Access Token，亦無需跨試算表讀取家長綁定！';

  ui.alert('⚙️ School LINE Hub 中樞設定', statusMsg, ui.ButtonSet.OK);
}

/**
 * 📢 發送【平時成績已更新】LINE 通知給已綁定家長（選單呼叫，支援全班或個別發送）
 * 透過 School LINE Hub 中樞發送座號、姓名與成績摘要
 */
function sendLineScoreUpdateNotificationMenu() {
  var ui = SpreadsheetApp.getUi();
  var status = getLinePushStatus_();

  // 詢問發送範圍：全班 vs 個別學生
  var askPrompt = ui.prompt(
    '📢 發送【平時成績已更新】LINE 通知（School LINE Hub 中樞）',
    '請選擇發送對象：\n' +
    '•【全班發送】：直接按「確定」（或輸入 ALL）\n' +
    '•【個別學生發送】：請輸入學生的「座號」（例如：5）或「查詢代碼」\n\n' +
    '目前全班已綁定：' + status.boundParentsCount + ' 位家長（涵蓋 ' + status.boundStudentsCount + ' 位學生）',
    ui.ButtonSet.OK_CANCEL
  );

  if (askPrompt.getSelectedButton() !== ui.Button.OK) return;

  var targetInput = askPrompt.getResponseText().trim().toUpperCase();
  var params = {};

  if (targetInput && targetInput !== 'ALL' && targetInput !== '全班') {
    var targetStudent = null;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    var rData = sSheet.getDataRange().getValues();
    for (var r = 1; r < rData.length; r++) {
      var sSeat = normalizeSeatNumber(rData[r][1]);
      var sTok = (rData[r][3] || '').toString().trim().toUpperCase();
      if (sSeat === targetInput || normalizeSeatNumber(targetInput) === sSeat || sTok === targetInput) {
        targetStudent = { seat: sSeat, name: rData[r][2], token: sTok };
        break;
      }
    }

    if (!targetStudent) {
      ui.alert('⚠️ 查無座號或代碼為「' + targetInput + '」的學生！');
      return;
    }

    params.token = targetStudent.token;

    var confirmSingleMsg = '📢【個別發送成績更新通知確認】\n\n' +
      '• 推播對象：' + targetStudent.seat + '號 ' + targetStudent.name + '\n' +
      '• 查詢代碼：' + targetStudent.token + '\n' +
      '• 推播管道：School LINE Hub 中樞\n' +
      '• 訊息內容：自動附帶各科最新成績摘要與個人專屬手札連結。\n\n' +
      '確定要透過中樞發送通知給該生家長嗎？';

    var askConfirm = ui.alert('📢 確認發送 LINE 通知', confirmSingleMsg, ui.ButtonSet.YES_NO);
    if (askConfirm !== ui.Button.YES) return;
  } else {
    var confirmMsg = '📢【全班發送成績更新通知確認】\n\n' +
      '• 全班學生人數：' + status.totalStudents + ' 人\n' +
      '• 已成功綁定 LINE 家長：' + status.boundParentsCount + ' 位（涵蓋 ' + status.boundStudentsCount + ' 位學生）\n' +
      '• 推播管道：School LINE Hub 中樞\n' +
      '• 訊息內容：自動附帶每位學生各自的成績摘要與個人專屬手札連結。\n\n' +
      '確定要立即透過中樞發送全班 LINE 成績更新通知嗎？';

    var askConfirm = ui.alert('📢 確認發送 LINE 通知', confirmMsg, ui.ButtonSet.YES_NO);
    if (askConfirm !== ui.Button.YES) return;
  }

  var result = executeLineScorePushNotifications_(params);
  if (result.success) {
    var resultMsg = '🎉【School LINE Hub 推播發送完畢】\n\n' +
      '• 成功發送：' + result.successCount + ' 則訊息\n' +
      (result.failCount > 0 ? ('• 發送失敗：' + result.failCount + ' 則（可能家長已封鎖官方帳號）\n') : '') +
      (result.skippedCount > 0 ? ('• 略過未綁定：' + result.skippedCount + ' 位學生\n') : '') +
      '• 涵蓋學生：' + result.studentCount + ' 位\n\n' +
      '家長已在 LINE 官方帳號收到成績摘要與手札連結！';
    ui.alert(resultMsg);
  } else {
    ui.alert('⚠️ 發送失敗：' + result.message);
  }
}

// 導出核心引擎供 Loader.gs 熱更新調用
var MPScoreDailyEngine = {
  setupStandardSheets: setupStandardSheets,
  createNewSubjectSheetFromRoster: createNewSubjectSheetFromRoster,
  syncRosterToSubjectSheets: syncRosterToSubjectSheets,
  syncStudentsFromSheet: syncStudentsFromSheet,
  generateTokensForStudents: generateTokensForStudents,
  previewClassStatistics: previewClassStatistics,
  autoFormatAndResizeAllSheets: autoFormatAndResizeAllSheets,
  openAccessLogsSheet: openAccessLogsSheet,
  clearAllAccessLogsMenu: clearAllAccessLogsMenu,
  openScoreUpdateLogsSheet: openScoreUpdateLogsSheet,
  clearScoreUpdateLogsMenu: clearScoreUpdateLogsMenu,
  convertAllSubjectSheetsToSchemeC: convertAllSubjectSheetsToSchemeC,
  manageTeacherAdminKeyMenu: manageTeacherAdminKeyMenu,
  setupLinePushSettingsMenu: setupLinePushSettingsMenu,
  sendLineScoreUpdateNotificationMenu: sendLineScoreUpdateNotificationMenu,
  getLinePushStatus: getLinePushStatus_,
  executeLineScorePushNotifications: executeLineScorePushNotifications_,
  getTeacherAdminKey: getTeacherAdminKey_,
  setTeacherAdminKey: setTeacherAdminKey_,
  getLineHubClient: getLineHubClient_,
  getLineHubConfig: getLineHubConfig_,
  clearAccessLogs: clearAccessLogs_,
  getRecentAccessLogs: getRecentAccessLogs_,
  getStudentAccessHistoryMap: getStudentAccessHistoryMap_,
  setupSchemeCColumnHeader: setupSchemeCColumnHeader_,
  clearSystemCache: clearSystemCache,
  onEdit: onEdit,
  handleDoGet: doGet,
  handleDoPost: doPost
};

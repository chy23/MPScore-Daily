/**
 * MPScore-Daily - 學生平時學習歷程手札 (Google Apps Script 後端核心)
 * 網站作者 / 系統規劃：楊家驊老師 (Chiahua Yang)
 *
 * 【核心架構 — 專為日常平時成績、習作、作業卷、語詞與單元測驗卷設計】：
 * 1. 【學生名冊】分頁：班級 | 座號 | 姓名 | 專屬代碼 | 啟用狀態 (可與 MPScore 段考系統 100% 共用代碼)
 * 2. 各科平時分頁（例如：國語平時、數學平時、社會平時、國語期中平時...）：
 *    - 支援「單列標題」（如：國習L1、作業卷L1、語詞L1、測驗卷L1）
 *    - 亦無痛支援老師原有的「雙列合併群組標題」（上列：國習 / 作業卷 / 語詞測驗 / 測驗卷；下列：L1、L2、L3...）
 *    - 自動略過「平均、總均、平時成績、名次」等統計欄，避免重複計入單項作業！
 * 3. 儲存格底色智慧辨識：
 *    - 【黃色底色】或填寫「缺/未交/待補」➔ 自動轉為前台溫馨提醒「⏳ 待補交 / 待補測」
 *    - 【粉紅/淺紅底色】或填寫「待訂正」➔ 自動轉為前台溫馨引導「🌱 需重點複習 / 待訂正」
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🌱 MPScore-Daily 平時工具')
    .addItem('🛠️ 建立初始表格（空白試算表第一次使用）', 'setupStandardSheets')
    .addSeparator()
    .addItem('➕ 新增平時科目分頁（自動帶入學生名單）', 'createNewSubjectSheetFromRoster')
    .addItem('🔄 用最新成績名單，更新「學生名冊」', 'syncStudentsFromSheet')
    .addItem('🔑 為全班產生「家長查詢代碼」', 'generateTokensForStudents')
    .addSeparator()
    .addItem('⏳ 查看全班「待補交・待訂正」彙整清單', 'previewMissingAndReviewItems')
    .addItem('📊 查看全班各科平時總均與四大支柱統計', 'previewClassStatistics')
    .addItem('📐 自動整理欄寬與表格排版', 'autoFormatAndResizeAllSheets')
    .addItem('🚀 立即更新雲端資料（改完分數或標色點此）', 'clearSystemCache')
    .addToUi();
}

/**
 * 清除系統雲端快取
 */
function clearSystemCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var cache = CacheService.getScriptCache();
    var keys = ['mpdaily_print_summary_v1'];
    for (var i = 1; i <= 60; i++) {
      var s = i < 10 ? '0' + i : '' + i;
      keys.push('mpdaily_st_v1_DEMO' + s);
      keys.push('mpdaily_st_v1_' + s);
    }
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sSheet = ss.getSheetByName('學生名冊');
      if (sSheet && sSheet.getLastRow() > 1) {
        var sData = sSheet.getDataRange().getValues();
        for (var r = 1; r < sData.length; r++) {
          var tok = (sData[r][3] || '').toString().trim();
          if (tok) keys.push('mpdaily_st_v1_' + tok);
        }
      }
    } catch (ignore) {}
    cache.removeAll(keys);
    ui.alert('✨ 雲端快取清除成功！', '家長端與查詢系統現已載入試算表最新平時成績、四大支柱統計與底色標記狀態。', ui.ButtonSet.OK);
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
 * 建立【學生名冊】、【國語平時】、【數學平時】、【社會平時】示範分頁，並含黃色/粉紅色儲存格標記示範
 */
function setupStandardSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. 建立或格式化【學生名冊】
  var sSheet = ss.getSheetByName('學生名冊') || ss.getSheets()[0];
  sSheet.setName('學生名冊');
  var sHeaders = [['班級', '座號', '姓名', '專屬代碼', '啟用狀態']];
  sSheet.getRange(1, 1, 1, 5).setValues(sHeaders)
    .setBackground('#b45309')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  var sampleRoster = [
    ['605', '01', '陳小明', 'DEMO01', true],
    ['605', '02', '陳小華', 'DEMO02', true],
    ['605', '03', '陳大美', 'DEMO03', true],
    ['605', '04', '陳志豪', 'DEMO04', true],
    ['605', '06', '黃小琪', 'DEMO06', true],
    ['605', '07', '李大偉', 'DEMO07', true],
    ['605', '08', '林子涵', 'DEMO08', true],
    ['605', '09', '凌雅婷', 'DEMO09', true],
    ['605', '10', '劉建銘', 'DEMO10', true],
    ['605', '11', '王冠宇', 'DEMO11', true],
    ['605', '12', '吳宗憲', 'DEMO12', true],
    ['605', '13', '嚴立婷', 'DEMO13', true],
    ['605', '14', '余承恩', 'DEMO14', true],
    ['605', '15', '嚴佳慧', 'DEMO15', true],
    ['605', '16', '羅品安', 'DEMO16', true],
    ['605', '17', '羅廷宇', 'DEMO17', true],
    ['605', '18', '蕭雅云', 'DEMO18', true],
    ['605', '19', '張凱傑', 'DEMO19', true],
    ['605', '20', '茅文傑', 'DEMO20', true],
    ['605', '21', '洪宇成', 'DEMO21', true],
    ['605', '22', '林郁潔', 'DEMO22', true],
    ['605', '23', '周奕翔', 'DEMO23', true],
    ['605', '24', '吳秉翰', 'DEMO24', true],
    ['605', '25', '李勝吉', 'DEMO25', true]
  ];
  if (sSheet.getLastRow() <= 1) {
    sSheet.getRange(2, 1, sampleRoster.length, 5).setValues(sampleRoster);
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
  sSheet.setFrozenRows(1);

  // 2. 建立【國語平時】分頁 (四大支柱：國習L1~L6、作業卷L1~L6、語詞L1~L6、測驗卷L1~L6)
  var chiHeaders = [
    '座號', '姓名',
    '國習L1', '國習L2', '國習L3', '國習L4', '國習L5', '國習L6',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5', '作業卷L6',
    '語詞L1', '語詞L2', '語詞L3', '語詞L4', '語詞L5', '語詞L6',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5', '測驗卷L6'
  ];
  var chiRows = sampleRoster.map(function(st, idx) {
    var base = idx % 5 === 0 ? 94 : (idx % 5 === 1 ? 88 : (idx % 5 === 2 ? 91 : (idx % 5 === 3 ? 78 : 84)));
    return [
      st[1], st[2],
      // 國習 L1~L6
      Math.min(100, base + 4), Math.min(100, base + 5), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 6), Math.min(100, base + 5),
      // 作業卷 L1~L6
      Math.min(100, base + 3), Math.min(100, base + 4), (idx === 1 ? '' : Math.min(100, base + 1)), Math.min(100, base + 5), Math.min(100, base + 3), Math.min(100, base + 4),
      // 語詞 L1~L6
      Math.min(100, base + 2), Math.min(100, base), Math.max(58, base - 8), Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 2),
      // 測驗卷 L1~L6 (讓 L3 呈現觀念斷層示範，L4~L6 回升)
      Math.min(100, base + 2), Math.min(100, base + 4), Math.max(56, base - 14), Math.max(64, base - 6), Math.min(100, base + 2), Math.min(100, base + 4)
    ];
  });
  var chiSheet = ss.getSheetByName('國語平時') || ss.insertSheet('國語平時');
  formatDailySubjectSheet(chiSheet, '115學年度第一學期 國語科平時學習歷程紀錄', chiHeaders, chiRows);
  // 標記示範：陳小華 (第4列, row=4) 的 作業卷L3 (col=11) 標為黃色待補交，測驗卷L3 (col=23) 標為粉紅色需重點複習
  chiSheet.getRange(4, 11).setBackground('#fef08a'); // 黃色：待補交
  chiSheet.getRange(4, 23).setBackground('#fecdd3'); // 粉紅：需重點複習
  chiSheet.getRange(6, 17).setValue('').setBackground('#fef08a'); // 陳志豪 語詞L3 待補測

  // 3. 建立【數學平時】分頁 (數習L1~L5、作業卷L1~L5、測驗卷L1~L5)
  var mathHeaders = [
    '座號', '姓名',
    '數習L1', '數習L2', '數習L3', '數習L4', '數習L5',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5'
  ];
  var mathRows = sampleRoster.map(function(st, idx) {
    var base = idx % 4 === 0 ? 90 : (idx % 4 === 1 ? 85 : (idx % 4 === 2 ? 93 : 76));
    return [
      st[1], st[2],
      Math.min(100, base + 5), Math.min(100, base + 4), Math.min(100, base + 2), Math.min(100, base + 6), Math.min(100, base + 5),
      Math.min(100, base + 2), Math.min(100, base + 3), Math.max(60, base - 6), Math.min(100, base + 4), Math.min(100, base + 3),
      Math.min(100, base + 3), Math.max(58, base - 12), Math.min(100, base), Math.min(100, base + 2), Math.min(100, base + 4)
    ];
  });
  var mathSheet = ss.getSheetByName('數學平時') || ss.insertSheet('數學平時');
  formatDailySubjectSheet(mathSheet, '115學年度第一學期 數學科平時學習歷程紀錄', mathHeaders, mathRows);
  mathSheet.getRange(4, 14).setBackground('#fecdd3'); // 陳小華 數學測驗卷L2 粉紅提醒

  // 4. 建立【社會平時】分頁 (社習L1~L4、作業卷L1~L4、測驗卷L1~L4)
  var socHeaders = [
    '座號', '姓名',
    '社習L1', '社習L2', '社習L3', '社習L4',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4'
  ];
  var socRows = sampleRoster.map(function(st, idx) {
    var base = idx % 4 === 0 ? 92 : (idx % 4 === 1 ? 88 : (idx % 4 === 2 ? 90 : 80));
    return [
      st[1], st[2],
      Math.min(100, base + 4), Math.min(100, base + 6), Math.min(100, base + 3), Math.min(100, base + 5),
      Math.min(100, base + 2), Math.min(100, base + 4), Math.min(100, base + 1), Math.min(100, base + 3),
      Math.min(100, base), Math.min(100, base + 2), Math.max(66, base - 8), Math.min(100, base + 4)
    ];
  });
  var socSheet = ss.getSheetByName('社會平時') || ss.insertSheet('社會平時');
  formatDailySubjectSheet(socSheet, '115學年度第一學期 社會科平時學習歷程紀錄', socHeaders, socRows);

  ui.alert(
    '🎉 平時學習歷程表格建置完成！\n\n' +
    '已為您建立：\n' +
    '1. 【學生名冊】(可與 MPScore 段考系統共用同一組學生代碼)\n' +
    '2. 【國語平時】(含國習、作業卷、語詞、測驗卷 L1~L6)\n' +
    '3. 【數學平時】(含數習、作業卷、測驗卷 L1~L5)\n' +
    '4. 【社會平時】(含社習、作業卷、測驗卷 L1~L4)\n\n' +
    '💡 貼心秘訣：\n' +
    '• 缺交或待補測的格子塗上【黃色底色】，家長端會自動顯示「⏳ 待補交 / 待補測」！\n' +
    '• 需加強或待訂正的格子塗上【粉紅色底色】，家長端會自動顯示「🌱 需重點複習」！'
  );
}

/**
 * 格式化平時成績科目分頁
 */
function formatDailySubjectSheet(sheet, titleText, headers, rowsData) {
  sheet.clear();
  var numCols = headers.length;

  // 第 1 列：大標題
  sheet.getRange(1, 1, 1, numCols).merge()
    .setValue(titleText)
    .setBackground('#78350f') // 溫潤深琥珀褐
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  // 第 2 列：欄位標題 (依四大支柱自動賦予柔和底色區隔)
  sheet.getRange(2, 1, 1, numCols).setValues([headers])
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  for (var c = 1; c <= numCols; c++) {
    var h = headers[c - 1] || '';
    var cell = sheet.getRange(2, c);
    if (c <= 2) {
      cell.setBackground('#334155').setFontColor('#ffffff');
    } else if (h.indexOf('習') !== -1) {
      cell.setBackground('#d97706').setFontColor('#ffffff'); // 課堂習作：暖陽琥珀
    } else if (h.indexOf('作業') !== -1 || h.indexOf('練習') !== -1) {
      cell.setBackground('#059669').setFontColor('#ffffff'); // 課後作業卷：萌芽翠綠
    } else if (h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1) {
      cell.setBackground('#0284c7').setFontColor('#ffffff'); // 語詞基礎：晴空青藍
    } else {
      cell.setBackground('#b45309').setFontColor('#ffffff'); // 單元測驗卷：深金琥珀
    }
  }

  // 第 3 列起：原始成績
  if (rowsData && rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, numCols).setValues(rowsData)
      .setHorizontalAlignment('center');
  }

  // 設定舒適欄寬與列高
  sheet.setColumnWidth(1, 70);  // 座號
  sheet.setColumnWidth(2, 110); // 姓名
  for (var col = 3; col <= numCols; col++) {
    sheet.setColumnWidth(col, 82);
  }
  sheet.setRowHeight(1, 38);
  sheet.setRowHeight(2, 32);
  var lastRow = Math.max(2, sheet.getLastRow());
  for (var r = 3; r <= lastRow; r++) {
    sheet.setRowHeight(r, 28);
  }
  sheet.getRange(1, 1, lastRow, numCols).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
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
      output = getStudentDailyPortfolioByToken(params.token);
    } else if (action === 'getStudentsForPrint') {
      output = getStudentListForTeacher();
    } else if (action === 'appendDailyScores') {
      var body = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = appendDailyScores(body);
    } else {
      output = { success: false, message: '未知的操作請求: ' + action };
    }
  } catch (err) {
    output = { success: false, message: '系統處理錯誤: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

/**
 * 判斷是否為「總表/彙整表」或需忽略之系統分頁
 */
function isSummaryOrSystemSheet(sheetName) {
  if (!sheetName) return true;
  if (sheetName === '學生名冊' || sheetName.startsWith('_')) return true;
  if (sheetName.indexOf('總表') !== -1 || sheetName.indexOf('總成績') !== -1) return true;
  return false;
}

/**
 * 判斷欄位標題是否為「統計結果欄」（例如：平均、總均、平時成績、名次、資源班），而非單次作業或測驗
 */
function isSummaryColumnHeader(headerText) {
  if (!headerText) return true;
  var h = headerText.toString().trim();
  if (!h) return true;
  if (h === '座號' || h === '姓名' || h === '班級' || h === '學號' || h === '備註') return true;
  var summaryKeywords = ['平均', '總均', '總分', '平時成績', '平時總成績', '學期成績', '期中平時', '期末平時', '排名', '名次', '資源班', '加權', '百分比', '占比'];
  for (var i = 0; i < summaryKeywords.length; i++) {
    if (h.indexOf(summaryKeywords[i]) !== -1) return true;
  }
  return false;
}

/**
 * 智慧判斷單項作業所屬之「四大支柱」分類
 * - workbook:   課堂習作表現 (含「習」)
 * - homework:   課後作業卷   (含「作業」、「練習」、「學習單」)
 * - foundation: 語詞基礎聽寫 (含「語詞」、「聽寫」、「圈詞」、「生字」、「單字」、「計算」)
 * - quiz:       單元測驗小考 (含「測驗」、「考」、「卷」、「單元」或預設單元評量)
 */
function classifyPillar(itemName) {
  var h = (itemName || '').toString().trim();
  if (h.indexOf('習') !== -1) {
    return { key: 'workbook', label: '課堂習作表現', shortLabel: '習作表現', icon: '📘' };
  }
  if (h.indexOf('作業') !== -1 || h.indexOf('練習') !== -1 || h.indexOf('學習單') !== -1) {
    return { key: 'homework', label: '課後作業卷', shortLabel: '作業卷', icon: '📝' };
  }
  if (h.indexOf('語詞') !== -1 || h.indexOf('聽寫') !== -1 || h.indexOf('圈詞') !== -1 || h.indexOf('生字') !== -1 || h.indexOf('單字') !== -1 || h.indexOf('計算') !== -1) {
    return { key: 'foundation', label: '語詞基礎聽寫', shortLabel: '語詞聽寫', icon: '🔤' };
  }
  return { key: 'quiz', label: '單元測驗小考', shortLabel: '單元測驗', icon: '🎯' };
}

/**
 * 從項目名稱中提取「單元標籤」(如 L1, L2, 第3課 -> L3, U1 -> L1)
 */
function extractUnitCode(itemName) {
  var h = (itemName || '').toString().trim();
  var m1 = h.match(/[LUlu單元課第]\s*0*(\d{1,2})/);
  if (m1 && m1[1]) {
    return 'L' + m1[1];
  }
  var m2 = h.match(/(\d{1,2})/);
  if (m2 && m2[1]) {
    return 'L' + m2[1];
  }
  return h;
}

/**
 * 智慧辨識儲存格背景顏色與數值狀態
 * - 黃色系 (#ffff00, #fff2cc, #fef08a, #fde047, #fce5cd 等) 或文字含「缺/未交/待補」➔ 'missing' (⏳ 待補交/待補測)
 * - 粉紅/淺紅系 (#f4cccc, #ea9999, #fecdd3, #fda4af 等) 或文字含「訂正」➔ 'review' (🌱 需重點複習/待訂正)
 */
function detectCellStatus(bgColor, rawVal, numScore, colHasAnyScores) {
  var bg = (bgColor || '#ffffff').toString().toLowerCase().trim();
  var strVal = (rawVal === null || rawVal === undefined) ? '' : rawVal.toString().trim();

  // 1. 先檢查文字關鍵字
  if (strVal.indexOf('缺') !== -1 || strVal.indexOf('未交') !== -1 || strVal.indexOf('待補') !== -1 || strVal.indexOf('補考') !== -1 || strVal === '?') {
    return { status: 'missing', statusLabel: '⏳ 待補交 / 待補測', colorTag: 'yellow' };
  }
  if (strVal.indexOf('訂正') !== -1 || strVal.indexOf('重寫') !== -1 || strVal.indexOf('加強') !== -1) {
    return { status: 'review', statusLabel: '🌱 需重點複習 / 待訂正', colorTag: 'pink' };
  }

  // 2. 解析 Hex RGB 判斷黃色或粉紅色底色
  if (/^#[0-9a-f]{6}$/.test(bg) && bg !== '#ffffff' && bg !== '#f8fafc' && bg !== '#f1f5f9') {
    var r = parseInt(bg.substr(1, 2), 16);
    var g = parseInt(bg.substr(3, 2), 16);
    var b = parseInt(bg.substr(5, 2), 16);

    // 黃色 / 暖橘黃標記 (R 高, G 高, B 明顯偏低)
    if (r >= 220 && g >= 195 && b <= 185 && (r - b >= 45) && (g - b >= 25)) {
      return {
        status: 'missing',
        statusLabel: numScore !== null ? '⏳ 補交/補測紀錄' : '⏳ 待補交 / 待補測',
        colorTag: 'yellow'
      };
    }

    // 粉紅 / 淺紅標記 (R 高, R 明顯大於 G)
    if (r >= 210 && g <= 210 && (r - g >= 25)) {
      return {
        status: 'review',
        statusLabel: '🌱 需重點複習 / 待訂正',
        colorTag: 'pink'
      };
    }
  }

  // 3. 若無底色但分數未滿 60 分，給予約定複習提醒
  if (numScore !== null && numScore < 60) {
    return { status: 'review', statusLabel: '🌱 建議加強複習', colorTag: 'pink' };
  }

  // 4. 若該欄全班多數人已有成績，但該生空白
  if (numScore === null && !strVal && colHasAnyScores) {
    return { status: 'missing', statusLabel: '⏳ 尚未登錄 / 待補交', colorTag: 'yellow' };
  }

  return { status: 'done', statusLabel: '✅ 已完成', colorTag: 'normal' };
}

/**
 * 解析單一平時成績工作表（自動支援單列標題 & 雙列群組標題 & 儲存格背景顏色讀取）
 */
function parseDailySubjectSheet(sheet) {
  var sheetName = sheet.getName();
  var range = sheet.getDataRange();
  var data = range.getValues();
  var backgrounds = range.getBackgrounds();

  if (data.length < 2) return null;

  // 尋找包含「座號」或「姓名」的主標題列
  var headerRowIdx = -1;
  for (var r = 0; r < Math.min(5, data.length); r++) {
    var rowStr = data[r].join(' ');
    if (rowStr.indexOf('座號') !== -1 || rowStr.indexOf('姓名') !== -1) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) headerRowIdx = 1;

  // 判斷是否為雙列標題（例如 headerRowIdx 列有「國習」「作業卷」，而 headerRowIdx+1 列是「L1」「L2」且沒有學生姓名）
  var groupRow = null;
  var itemHeaderRow = data[headerRowIdx];
  var dataStartRowIdx = headerRowIdx + 1;

  if (headerRowIdx > 0) {
    // 檢查上一列是否為群組標題列（非單純全表合併大標題）
    var prevRow = data[headerRowIdx - 1];
    var nonEmptyPrev = 0;
    for (var pc = 0; pc < prevRow.length; pc++) {
      if ((prevRow[pc] || '').toString().trim()) nonEmptyPrev++;
    }
    if (nonEmptyPrev >= 2) {
      groupRow = prevRow;
    }
  }

  // 檢查下一列是否為子標題列 (例如座號姓名下方第一列在座號欄是空白，但後面是 L1, L2...)
  if (headerRowIdx + 1 < data.length) {
    var nextRow = data[headerRowIdx + 1];
    var nextRowSeat = (nextRow[0] || '').toString().trim();
    var nextRowName = (nextRow[1] || '').toString().trim();
    var nextRowJoined = nextRow.slice(2).join(' ');
    if (!nextRowSeat && !nextRowName && (nextRowJoined.indexOf('L1') !== -1 || nextRowJoined.indexOf('1') !== -1)) {
      groupRow = data[headerRowIdx];
      itemHeaderRow = nextRow;
      dataStartRowIdx = headerRowIdx + 2;
    }
  }

  // 取分頁顯示標題
  var sheetTitle = sheetName;
  if (headerRowIdx >= 1 && data[0][0]) {
    var firstCell = data[0][0].toString().trim();
    if (firstCell && firstCell.indexOf('座號') === -1) {
      sheetTitle = firstCell;
    }
  }

  // 找出座號與姓名欄
  var seatCol = 0;
  var nameCol = 1;
  for (var c = 0; c < itemHeaderRow.length; c++) {
    var hText = (itemHeaderRow[c] || (groupRow ? groupRow[c] : '') || '').toString().trim();
    if (hText.indexOf('座號') !== -1) seatCol = c;
    if (hText.indexOf('姓名') !== -1) nameCol = c;
  }

  // 建立有效評量項目欄位清單 (自動組裝群組名稱 + 子標題名稱，並過濾平均/總均統計欄)
  var activeGroupLabel = '';
  var itemCols = [];

  for (var col = 0; col < itemHeaderRow.length; col++) {
    if (col === seatCol || col === nameCol) continue;

    if (groupRow) {
      var gCell = (groupRow[col] || '').toString().trim();
      if (gCell && gCell.indexOf('座號') === -1 && gCell.indexOf('姓名') === -1) {
        activeGroupLabel = gCell;
      }
    }

    var rawItemTitle = (itemHeaderRow[col] || '').toString().trim();
    if (!rawItemTitle && !activeGroupLabel) continue;
    if (isSummaryColumnHeader(rawItemTitle) || isSummaryColumnHeader(activeGroupLabel)) continue;

    var fullItemTitle = rawItemTitle;
    // 若子標題僅為 L1, L2, 1, 2 等簡碼，且有上層群組名稱，則組合為「國習L1」
    if (activeGroupLabel && (/^[LUlu]?\d{1,2}$/.test(rawItemTitle) || rawItemTitle.length <= 3) && rawItemTitle.indexOf(activeGroupLabel) === -1) {
      fullItemTitle = activeGroupLabel + rawItemTitle;
    } else if (!rawItemTitle && activeGroupLabel) {
      fullItemTitle = activeGroupLabel;
    }

    var pillarInfo = classifyPillar(fullItemTitle);
    var unitCode = extractUnitCode(fullItemTitle);

    itemCols.push({
      colIdx: col,
      title: fullItemTitle,
      shortTitle: rawItemTitle || fullItemTitle,
      pillarKey: pillarInfo.key,
      pillarLabel: pillarInfo.label,
      pillarShort: pillarInfo.shortLabel,
      unitCode: unitCode
    });
  }

  // 第一階段：掃描全班學生各欄成績，計算每項作業的班級平均與是否有登錄成績
  var rawStudents = [];
  var colSums = {};
  var colCounts = {};

  for (var i = 0; i < itemCols.length; i++) {
    colSums[itemCols[i].colIdx] = 0;
    colCounts[itemCols[i].colIdx] = 0;
  }

  for (var rIdx = dataStartRowIdx; rIdx < data.length; rIdx++) {
    var row = data[rIdx];
    var bgRow = backgrounds[rIdx] || [];
    var seatVal = row[seatCol] !== undefined ? row[seatCol].toString().trim() : '';
    var nameVal = row[nameCol] !== undefined ? row[nameCol].toString().trim() : '';

    if (!seatVal && !nameVal) continue;
    if (seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1 || nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal);
    rawStudents.push({
      rowIdx: rIdx,
      seat: normSeat,
      name: nameVal,
      row: row,
      bgRow: bgRow
    });

    for (var j = 0; j < itemCols.length; j++) {
      var ic = itemCols[j];
      var cellVal = row[ic.colIdx];
      var num = (cellVal !== '' && cellVal !== null && !isNaN(Number(cellVal))) ? Number(cellVal) : null;
      if (num !== null) {
        colSums[ic.colIdx] += num;
        colCounts[ic.colIdx] += 1;
      }
    }
  }

  var colClassAvgs = {};
  for (var k = 0; k < itemCols.length; k++) {
    var cIdx = itemCols[k].colIdx;
    colClassAvgs[cIdx] = colCounts[cIdx] > 0 ? Number((colSums[cIdx] / colCounts[cIdx]).toFixed(1)) : null;
  }

  // 第二階段：為每位學生計算四大支柱平均、平時總平均、單元時序折線與待補交/待訂正清單
  var studentMap = {};
  var classOverallSum = 0;
  var classOverallCount = 0;
  var pillarClassSums = { workbook: 0, homework: 0, foundation: 0, quiz: 0 };
  var pillarClassCounts = { workbook: 0, homework: 0, foundation: 0, quiz: 0 };

  for (var s = 0; s < rawStudents.length; s++) {
    var st = rawStudents[s];
    var items = [];
    var pillarBuckets = {
      workbook: { scores: [], totalItems: 0, completedItems: 0, missingCount: 0, reviewCount: 0 },
      homework: { scores: [], totalItems: 0, completedItems: 0, missingCount: 0, reviewCount: 0 },
      foundation: { scores: [], totalItems: 0, completedItems: 0, missingCount: 0, reviewCount: 0 },
      quiz: { scores: [], totalItems: 0, completedItems: 0, missingCount: 0, reviewCount: 0 }
    };

    var allValidScores = [];
    var missingItems = [];
    var reviewItems = [];
    var unitMap = {};
    var unitOrder = [];

    for (var m = 0; m < itemCols.length; m++) {
      var colMeta = itemCols[m];
      var rawVal = st.row[colMeta.colIdx];
      var bgHex = st.bgRow[colMeta.colIdx] || '#ffffff';
      var hasScoresInCol = colCounts[colMeta.colIdx] > 0;

      // 若整欄全班都還沒登記且無底色，則略過不視為缺交
      if (!hasScoresInCol && (rawVal === '' || rawVal === null) && bgHex === '#ffffff') {
        continue;
      }

      var numScore = (rawVal !== '' && rawVal !== null && !isNaN(Number(rawVal))) ? Number(rawVal) : null;
      var stStatus = detectCellStatus(bgHex, rawVal, numScore, hasScoresInCol);
      var classAvg = colClassAvgs[colMeta.colIdx];

      var bkt = pillarBuckets[colMeta.pillarKey] || pillarBuckets.quiz;
      bkt.totalItems += 1;

      if (numScore !== null) {
        bkt.scores.push(numScore);
        bkt.completedItems += 1;
        allValidScores.push(numScore);
      }

      if (stStatus.status === 'missing') {
        bkt.missingCount += 1;
        missingItems.push({
          title: colMeta.title,
          pillarLabel: colMeta.pillarShort,
          statusLabel: stStatus.statusLabel,
          score: numScore
        });
      } else if (stStatus.status === 'review') {
        bkt.reviewCount += 1;
        reviewItems.push({
          title: colMeta.title,
          pillarLabel: colMeta.pillarShort,
          statusLabel: stStatus.statusLabel,
          score: numScore
        });
      }

      // 單元脈絡歸納 (L1, L2, L3...)
      var uKey = colMeta.unitCode;
      if (uKey) {
        if (!unitMap[uKey]) {
          unitMap[uKey] = {
            unit: uKey,
            scores: [],
            classAvgs: [],
            byPillar: {}
          };
          unitOrder.push(uKey);
        }
        if (numScore !== null) {
          unitMap[uKey].scores.push(numScore);
          unitMap[uKey].byPillar[colMeta.pillarKey] = numScore;
        }
        if (classAvg !== null) {
          unitMap[uKey].classAvgs.push(classAvg);
          unitMap[uKey].byPillar[colMeta.pillarKey + '_classAvg'] = classAvg;
        }
      }

      items.push({
        title: colMeta.title,
        shortTitle: colMeta.shortTitle,
        pillarKey: colMeta.pillarKey,
        pillarLabel: colMeta.pillarLabel,
        pillarShort: colMeta.pillarShort,
        unitCode: colMeta.unitCode,
        score: numScore,
        rawText: rawVal !== null && rawVal !== undefined ? rawVal.toString() : '',
        classAvg: classAvg,
        status: stStatus.status,
        statusLabel: stStatus.statusLabel,
        colorTag: stStatus.colorTag
      });
    }

    // 計算四大支柱平均
    var pillarsSummary = {};
    var pillarKeys = ['workbook', 'homework', 'foundation', 'quiz'];
    var pillarMetaMap = {
      workbook: { label: '課堂習作表現', shortLabel: '習作表現', icon: '📘' },
      homework: { label: '課後作業卷', shortLabel: '作業卷', icon: '📝' },
      foundation: { label: '語詞基礎聽寫', shortLabel: '語詞聽寫', icon: '🔤' },
      quiz: { label: '單元測驗小考', shortLabel: '單元測驗', icon: '🎯' }
    };

    var weightedPillarAvgs = [];
    for (var pkIdx = 0; pkIdx < pillarKeys.length; pkIdx++) {
      var pk = pillarKeys[pkIdx];
      var pb = pillarBuckets[pk];
      var pAvg = null;
      if (pb.scores.length > 0) {
        var pSum = 0;
        for (var ps = 0; ps < pb.scores.length; ps++) pSum += pb.scores[ps];
        pAvg = Number((pSum / pb.scores.length).toFixed(1));
        weightedPillarAvgs.push(pAvg);
        pillarClassSums[pk] += pAvg;
        pillarClassCounts[pk] += 1;
      }
      pillarsSummary[pk] = {
        key: pk,
        label: pillarMetaMap[pk].label,
        shortLabel: pillarMetaMap[pk].shortLabel,
        icon: pillarMetaMap[pk].icon,
        average: pAvg,
        totalItems: pb.totalItems,
        completedItems: pb.completedItems,
        missingCount: pb.missingCount,
        reviewCount: pb.reviewCount,
        badge: getPillarFeedbackBadge(pAvg, pb.missingCount)
      };
    }

    // 該科平時總平均（若有多個支柱則採各支柱均權或所有有效項目平均）
    var overallAvg = null;
    if (allValidScores.length > 0) {
      var totalS = 0;
      for (var as = 0; as < allValidScores.length; as++) totalS += allValidScores[as];
      overallAvg = Number((totalS / allValidScores.length).toFixed(1));
      classOverallSum += overallAvg;
      classOverallCount += 1;
    }

    // 整理單元折線圖陣列
    var unitTimeline = unitOrder.map(function(u) {
      var uObj = unitMap[u];
      var uAvg = null;
      if (uObj.scores.length > 0) {
        var us = 0;
        for (var ui = 0; ui < uObj.scores.length; ui++) us += uObj.scores[ui];
        uAvg = Number((us / uObj.scores.length).toFixed(1));
      }
      var cAvg = null;
      if (uObj.classAvgs.length > 0) {
        var cs = 0;
        for (var ci = 0; ci < uObj.classAvgs.length; ci++) cs += uObj.classAvgs[ci];
        cAvg = Number((cs / uObj.classAvgs.length).toFixed(1));
      }
      return {
        unit: u,
        overallScore: uAvg,
        classAvg: cAvg,
        workbook: uObj.byPillar.workbook !== undefined ? uObj.byPillar.workbook : null,
        homework: uObj.byPillar.homework !== undefined ? uObj.byPillar.homework : null,
        foundation: uObj.byPillar.foundation !== undefined ? uObj.byPillar.foundation : null,
        quiz: uObj.byPillar.quiz !== undefined ? uObj.byPillar.quiz : null,
        quizClassAvg: uObj.byPillar.quiz_classAvg !== undefined ? uObj.byPillar.quiz_classAvg : cAvg
      };
    });

    studentMap[st.seat] = {
      sheetName: sheetName,
      subjectTitle: sheetTitle,
      subjectShortName: inferSubjectShortName(sheetName, sheetTitle),
      overallAverage: overallAvg,
      totalItemsCount: items.length,
      completedItemsCount: allValidScores.length,
      missingCount: missingItems.length,
      reviewCount: reviewItems.length,
      missingItems: missingItems,
      reviewItems: reviewItems,
      pillars: pillarsSummary,
      unitTimeline: unitTimeline,
      items: items
    };
  }

  var subjectClassAvg = classOverallCount > 0 ? Number((classOverallSum / classOverallCount).toFixed(1)) : 85.0;
  var pillarClassAvgs = {};
  ['workbook', 'homework', 'foundation', 'quiz'].forEach(function(pk) {
    pillarClassAvgs[pk] = pillarClassCounts[pk] > 0 ? Number((pillarClassSums[pk] / pillarClassCounts[pk]).toFixed(1)) : subjectClassAvg;
  });

  // 補入班級平均與教師溫馨評語
  Object.keys(studentMap).forEach(function(seatKey) {
    var stRecord = studentMap[seatKey];
    stRecord.classOverallAverage = subjectClassAvg;
    ['workbook', 'homework', 'foundation', 'quiz'].forEach(function(pk) {
      if (stRecord.pillars[pk]) {
        stRecord.pillars[pk].classAverage = pillarClassAvgs[pk];
      }
    });
    stRecord.teacherAdvice = buildDailyWarmAdvice(stRecord);
  });

  return {
    sheetName: sheetName,
    subjectTitle: sheetTitle,
    subjectShortName: inferSubjectShortName(sheetName, sheetTitle),
    classOverallAverage: subjectClassAvg,
    studentMap: studentMap
  };
}

/**
 * 推斷簡短科目名稱 (例如：「國語平時」->「國語」)
 */
function inferSubjectShortName(sheetName, sheetTitle) {
  var combined = (sheetName || '') + ' ' + (sheetTitle || '');
  if (combined.indexOf('國語') !== -1 || combined.indexOf('國文') !== -1) return '國語';
  if (combined.indexOf('數學') !== -1) return '數學';
  if (combined.indexOf('社會') !== -1) return '社會';
  if (combined.indexOf('自然') !== -1) return '自然';
  if (combined.indexOf('英文') !== -1 || combined.indexOf('英語') !== -1) return '英文';
  return (sheetName || '平時').replace(/平時成績|平時/g, '').trim() || sheetName;
}

/**
 * 根據各支柱分數與缺交狀況給予溫暖鼓勵徽章
 */
function getPillarFeedbackBadge(avg, missingCount) {
  if (missingCount > 0) return '⏳ 待補齊進度';
  if (avg === null || avg === undefined) return '🌱 持續累積中';
  if (avg >= 95) return '表現卓越 🌟';
  if (avg >= 90) return '認真扎實 ⭐';
  if (avg >= 85) return '穩步精進 🌿';
  if (avg >= 75) return '觀念良好 💪';
  return '潛力開發 🌱';
}

/**
 * 產生個人化平時學習歷程溫馨叮嚀
 */
function buildDailyWarmAdvice(stRecord) {
  var sub = stRecord.subjectShortName || '本科';
  var parts = [];

  var wb = stRecord.pillars && stRecord.pillars.workbook ? stRecord.pillars.workbook.average : null;
  var hw = stRecord.pillars && stRecord.pillars.homework ? stRecord.pillars.homework.average : null;
  var qz = stRecord.pillars && stRecord.pillars.quiz ? stRecord.pillars.quiz.average : null;

  if ((wb !== null && wb >= 90) || (hw !== null && hw >= 90)) {
    parts.push('孩子在' + sub + '的日常習作與課後作業態度非常認真扎實，展現了良好的課堂專注度！');
  } else if (wb !== null || hw !== null) {
    parts.push('孩子持續累積' + sub + '的日常作業經驗，建議每天完成作業後花 5 分鐘自主檢查訂正，能讓基礎更穩固。');
  }

  // 尋找單元折線圖中的最高與最弱單元
  if (stRecord.unitTimeline && stRecord.unitTimeline.length >= 2) {
    var lowestUnit = null;
    var highestUnit = null;
    stRecord.unitTimeline.forEach(function(u) {
      var val = u.quiz !== null ? u.quiz : u.overallScore;
      if (val !== null) {
        if (!lowestUnit || val < lowestUnit.val) lowestUnit = { unit: u.unit, val: val };
        if (!highestUnit || val > highestUnit.val) highestUnit = { unit: u.unit, val: val };
      }
    });
    if (highestUnit && lowestUnit && highestUnit.unit !== lowestUnit.unit && (highestUnit.val - lowestUnit.val >= 10)) {
      parts.push('從單元脈絡觀察，【' + highestUnit.unit + '】（' + highestUnit.val + '分）吸收得非常棒；而【' + lowestUnit.unit + '】（' + lowestUnit.val + '分）相對遇到些許瓶頸，建議可針對該單元的錯題重點回顧，及早補強觀念斷層。');
    } else if (qz !== null && qz >= 88) {
      parts.push('各單元測驗小考表現相當穩定，觀念理解透徹，請持續保持這份好習慣！');
    }
  }

  if (stRecord.missingCount > 0) {
    var mNames = stRecord.missingItems.map(function(m) { return m.title; }).join('、');
    parts.push('溫馨提醒：目前尚有 ' + stRecord.missingCount + ' 項平時紀錄（' + mNames + '）待補交或補測，請家長陪伴孩子於近期完成並交回給老師喔！');
  } else if (stRecord.reviewCount > 0) {
    var rNames = stRecord.reviewItems.map(function(r) { return r.title; }).join('、');
    parts.push('溫馨提醒：【' + rNames + '】老師已標記為重點複習/待訂正項目，弄懂錯題就是進步最快的捷徑！');
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
    if (!isSummaryOrSystemSheet(sName)) {
      subjectNames.push(sName);
    }
  }

  if (!sSheet) {
    return {
      success: false,
      message: '找不到【學生名冊】分頁！請在試算表上方點選【🌱 MPScore-Daily 平時工具 > 🛠️ 建立初始表格】。'
    };
  }

  var studentCount = Math.max(0, sSheet.getLastRow() - 1);
  return {
    success: true,
    message: '連線成功！已載入 ' + studentCount + ' 位學生名冊與 ' + subjectNames.length + ' 個平時成績分頁。',
    spreadsheetName: ss.getName(),
    studentCount: studentCount,
    subjectList: subjectNames
  };
}

/**
 * 家長端依 Token 查詢該生所有平時學習歷程資料 (高資安隔離，絕不外洩他生分數)
 */
function getStudentDailyPortfolioByToken(token) {
  var cleanToken = (token || '').toString().trim().toUpperCase();
  if (!cleanToken) {
    return { success: false, message: '請輸入或掃描有效的專屬查詢代碼！' };
  }

  var cache = CacheService.getScriptCache();
  var cacheKey = 'mpdaily_st_v1_' + cleanToken;
  var cachedJson = cache.get(cacheKey);
  if (cachedJson) {
    try {
      return JSON.parse(cachedJson);
    } catch (ignore) {}
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '系統尚未初始化【學生名冊】分頁。' };
  }

  var sData = sSheet.getDataRange().getValues();
  var matchedStudent = null;

  for (var r = 1; r < sData.length; r++) {
    var rowTok = (sData[r][3] || '').toString().trim().toUpperCase();
    var enabled = sData[r][4] !== false && sData[r][4].toString().toUpperCase() !== 'FALSE';
    if (rowTok === cleanToken && enabled) {
      var cls = (sData[r][0] || '605').toString().trim();
      var seat = normalizeSeatNumber(sData[r][1]);
      matchedStudent = {
        className: cls,
        seatNumber: seat,
        name: (sData[r][2] || '').toString().trim(),
        token: (sData[r][3] || '').toString().trim(),
        studentId: cls + '-' + seat
      };
      break;
    }
  }

  if (!matchedStudent) {
    return { success: false, message: '查無此專屬代碼或尚未啟用，請確認通知單上的代碼是否正確。' };
  }

  var sheets = ss.getSheets();
  var subjectsData = [];

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSummaryOrSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (parsed && parsed.studentMap[matchedStudent.seatNumber]) {
      subjectsData.push(parsed.studentMap[matchedStudent.seatNumber]);
    }
  }

  var responseObj = {
    success: true,
    student: matchedStudent,
    subjects: subjectsData,
    updatedAt: new Date().toISOString()
  };

  try {
    cache.put(cacheKey, JSON.stringify(responseObj), 600);
  } catch (ignore) {}

  return responseObj;
}

/**
 * 教師端同步名冊與全班各科平時統計摘要（供 A4 通知小卡 PDF 預覽與手機沙盒模擬使用）
 */
function getStudentListForTeacher() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '找不到【學生名冊】分頁！請先建立初始表格。' };
  }

  var sData = sSheet.getDataRange().getValues();
  var students = [];
  for (var r = 1; r < sData.length; r++) {
    var cls = (sData[r][0] || '605').toString().trim();
    var seat = normalizeSeatNumber(sData[r][1]);
    var name = (sData[r][2] || '').toString().trim();
    var token = (sData[r][3] || '').toString().trim();
    if (!seat && !name) continue;
    students.push({
      className: cls,
      seatNumber: seat,
      name: name,
      token: token,
      studentId: cls + '-' + seat,
      subjectStats: {}
    });
  }

  var sheets = ss.getSheets();
  var subjectList = [];

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSummaryOrSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (!parsed) continue;
    subjectList.push(parsed.sheetName);

    for (var s = 0; s < students.length; s++) {
      var st = students[s];
      var stSub = parsed.studentMap[st.seatNumber];
      if (stSub) {
        st.subjectStats[parsed.sheetName] = stSub;
      }
    }
  }

  return {
    success: true,
    students: students,
    subjectList: subjectList
  };
}

/**
 * 🔄 用最新成績名單，更新「學生名冊」
 * 保留老師設定的班級名稱（不強制重置）與既有專屬代碼，並維持舒適欄寬
 */
function syncStudentsFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    sSheet = ss.insertSheet('學生名冊', 0);
    sSheet.getRange(1, 1, 1, 5).setValues([['班級', '座號', '姓名', '專屬代碼', '啟用狀態']])
      .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  var sheets = ss.getSheets();
  var targetSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (!isSummaryOrSystemSheet(sheets[i].getName())) {
      targetSheet = sheets[i];
      break;
    }
  }

  if (!targetSheet) {
    ui.alert('請先建立至少一個平時成績分頁（例如：國語平時）並填入座號與姓名！');
    return;
  }

  // 讀取既有班級與代碼，避免覆蓋老師已修改的班級
  var existingBySeat = {};
  var fallbackClass = '605';
  var sData = sSheet.getDataRange().getValues();
  for (var s = 1; s < sData.length; s++) {
    var exCls = (sData[s][0] || '').toString().trim();
    var exSeat = normalizeSeatNumber(sData[s][1]);
    var exTok = (sData[s][3] || '').toString().trim();
    if (exCls) fallbackClass = exCls;
    if (exSeat) {
      existingBySeat[exSeat] = { className: exCls || fallbackClass, token: exTok };
    }
  }

  var data = targetSheet.getDataRange().getValues();
  var headerRowIdx = 1;
  for (var r = 0; r < Math.min(4, data.length); r++) {
    if (data[r].join(' ').indexOf('座號') !== -1 || data[r].join(' ').indexOf('姓名') !== -1) {
      headerRowIdx = r;
      break;
    }
  }

  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  var newRoster = [];

  for (var rowIdx = headerRowIdx + 1; rowIdx < data.length; rowIdx++) {
    var seatVal = data[rowIdx][0] !== undefined ? data[rowIdx][0].toString().trim() : '';
    var nameVal = data[rowIdx][1] !== undefined ? data[rowIdx][1].toString().trim() : '';
    if (!seatVal && !nameVal) continue;
    if (seatVal.indexOf('平均') !== -1 || nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal);
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

  if (sSheet.getLastRow() > 1) {
    sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 5).clearContent();
  }
  sSheet.getRange(2, 1, newRoster.length, 5).setValues(newRoster);

  // 固定舒適欄寬，絕不因 autoResizeColumns 造成文字擠壓
  sSheet.setColumnWidth(1, 90);
  sSheet.setColumnWidth(2, 80);
  sSheet.setColumnWidth(3, 120);
  sSheet.setColumnWidth(4, 130);
  sSheet.setColumnWidth(5, 100);

  ui.alert('🎉 同步完成！已從【' + targetSheet.getName() + '】更新 ' + newRoster.length + ' 位學生名冊，並完整保留班級名稱與專屬代碼！');
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
 * ➕ 新增平時科目分頁（自動帶入學生名單與四大支柱預設欄位）
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
    '➕ 新增平時科目分頁',
    '請輸入科目分頁名稱（例如：自然平時、英文平時、國語期末平時）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (promptResult.getSelectedButton() !== ui.Button.OK) return;

  var sheetTitle = promptResult.getResponseText().trim();
  if (!sheetTitle) return;
  if (ss.getSheetByName(sheetTitle)) {
    ui.alert('分頁【' + sheetTitle + '】已經存在！');
    return;
  }

  var subShort = inferSubjectShortName(sheetTitle, sheetTitle).charAt(0);
  var headers = [
    '座號', '姓名',
    subShort + '習L1', subShort + '習L2', subShort + '習L3', subShort + '習L4', subShort + '習L5',
    '作業卷L1', '作業卷L2', '作業卷L3', '作業卷L4', '作業卷L5',
    '測驗卷L1', '測驗卷L2', '測驗卷L3', '測驗卷L4', '測驗卷L5'
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
  formatDailySubjectSheet(newSheet, sheetTitle + ' 學習歷程紀錄', headers, rows);
  ui.alert('🎉 已成功建立【' + sheetTitle + '】並自動帶入全班 ' + rows.length + ' 位學生名單！');
}

/**
 * ⏳ 查看全班「待補交・待訂正」彙整清單
 */
function previewMissingAndReviewItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var res = getStudentListForTeacher();
  if (!res.success) {
    ui.alert(res.message);
    return;
  }

  var lines = [];
  res.students.forEach(function(st) {
    var stMissing = [];
    var stReview = [];
    Object.keys(st.subjectStats || {}).forEach(function(subKey) {
      var subObj = st.subjectStats[subKey];
      (subObj.missingItems || []).forEach(function(m) {
        stMissing.push(subObj.subjectShortName + ':' + m.title);
      });
      (subObj.reviewItems || []).forEach(function(rv) {
        stReview.push(subObj.subjectShortName + ':' + rv.title);
      });
    });
    if (stMissing.length > 0 || stReview.length > 0) {
      var line = st.seatNumber + '號 ' + st.name + ' ── ';
      if (stMissing.length > 0) line += '【⏳待補交: ' + stMissing.join('、') + '】 ';
      if (stReview.length > 0) line += '【🌱需複習: ' + stReview.join('、') + '】';
      lines.push(line);
    }
  });

  if (lines.length === 0) {
    ui.alert('🎉 太棒了！目前全班所有平時作業與測驗皆已完成，無黃色缺交或粉紅待訂正項目！');
  } else {
    ui.alert('⏳ 全班【待補交 / 待訂正】溫馨追蹤清單', lines.slice(0, 25).join('\n'), ui.ButtonSet.OK);
  }
}

/**
 * 📊 查看全班各科平時總均與四大支柱統計
 */
function previewClassStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheets = ss.getSheets();
  var summaryLines = [];

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (isSummaryOrSystemSheet(sh.getName())) continue;
    var parsed = parseDailySubjectSheet(sh);
    if (parsed) {
      summaryLines.push('📘 【' + parsed.sheetName + '】全班平時總平均：' + parsed.classOverallAverage + ' 分');
    }
  }

  if (summaryLines.length === 0) {
    ui.alert('尚未偵測到平時成績分頁！');
  } else {
    ui.alert('📊 全班各科平時學習歷程統計', summaryLines.join('\n\n'), ui.ButtonSet.OK);
  }
}

/**
 * 📐 自動整理欄寬與表格排版
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
      var lr = Math.max(2, sh.getLastRow());
      sh.getRange(1, 1, lr, 5).setVerticalAlignment('middle').setHorizontalAlignment('center');
      sh.setFrozenRows(1);
      count++;
    } else {
      var maxCol = Math.max(2, sh.getLastColumn());
      var maxRow = Math.max(2, sh.getLastRow());
      sh.setColumnWidth(1, 70);
      sh.setColumnWidth(2, 110);
      for (var c = 3; c <= maxCol; c++) {
        sh.setColumnWidth(c, 82);
      }
      sh.getRange(1, 1, maxRow, maxCol).setVerticalAlignment('middle').setHorizontalAlignment('center');
      sh.setFrozenRows(2);
      sh.setFrozenColumns(2);
      count++;
    }
  }

  SpreadsheetApp.getUi().alert('🎉 排版完成！已為您將 ' + count + ' 個分頁的欄寬、列高與凍結窗格最佳化！');
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
    return { success: false, message: '沒有平時成績資料可儲存' };
  }

  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  formatDailySubjectSheet(sheet, sheetName + ' 學習歷程紀錄', headers, rows);

  return {
    success: true,
    message: '成功建立並更新【' + sheetName + '】，共 ' + rows.length + ' 位學生平時紀錄！',
    sheetName: sheetName,
    studentCount: rows.length
  };
}

// 導出核心引擎供 Loader.gs 熱更新調用
var MPScoreDailyEngine = {
  setupStandardSheets: setupStandardSheets,
  createNewSubjectSheetFromRoster: createNewSubjectSheetFromRoster,
  syncStudentsFromSheet: syncStudentsFromSheet,
  generateTokensForStudents: generateTokensForStudents,
  previewMissingAndReviewItems: previewMissingAndReviewItems,
  previewClassStatistics: previewClassStatistics,
  autoFormatAndResizeAllSheets: autoFormatAndResizeAllSheets,
  clearSystemCache: clearSystemCache,
  handleDoGet: doGet,
  handleDoPost: doPost
};

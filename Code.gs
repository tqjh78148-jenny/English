/**
 * 🏫 頭前國中 907 班 英文單字小考 — 跨載具雲端同步系統 (Google Apps Script 後端)
 *
 * 特色：
 * 1. 學生在不同手機、平板、電腦打開 Google 網址，每次進去隨機重抽 10 題。
 * 2. 交卷後即時同步登記至 Google 雲端試算表。
 * 3. 任何同學在不同載具皆可開啟「907 班成績總覽」查看全班 01~31 號所有同學的成績與排名。
 * 4. 支援 Web App 介面與跨載具 JSON API (doGet / doPost)。
 */

function doGet(e) {
  // 1. 若外部載具/GitHub Pages API 查詢成績總覽
  if (e && e.parameter && (e.parameter.action === 'getAllScores' || e.parameter.action === 'getScores')) {
    const list = getAllScores();
    const jsonStr = JSON.stringify({ success: true, scores: list });
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 2. 若查詢目前老師鎖定的測驗起訖日期
  if (e && e.parameter && e.parameter.action === 'getRange') {
    const range = getActiveQuizRange();
    const jsonStr = JSON.stringify(range);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 3. 支援 GitHub Pages 透過 GET/JSONP 送出交卷成績（跨網域 100% 成功，絕不被瀏覽器擋）
  if (e && e.parameter && e.parameter.action === 'submitScore') {
    const payload = {
      className: e.parameter.className || '907',
      seatNo: e.parameter.seatNo,
      name: e.parameter.name,
      score: e.parameter.score,
      range: e.parameter.range,
      wrongWords: e.parameter.wrongWords
    };
    const result = submitStudentScore(payload);
    const jsonStr = JSON.stringify(result);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 4. 支援 GitHub Pages 清除成績（需驗證密碼 123456）
  if (e && e.parameter && e.parameter.action === 'clearScores') {
    const result = clearAllScores(e.parameter.password);
    const jsonStr = JSON.stringify(result);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 5. 支援 GitHub Pages 設定測驗日期範圍（需驗證密碼 123456）
  if (e && e.parameter && e.parameter.action === 'setRange') {
    const result = setActiveQuizRange(e.parameter.password, e.parameter.start, e.parameter.end);
    const jsonStr = JSON.stringify(result);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 6. 預設返回網頁介面（若未在 GAS 建立 index 檔案，則提示 API 正常運作）
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('頭前國中 907 班 英文聯絡簿單字小考')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (htmlErr) {
    return ContentService.createTextOutput("✅ Google Apps Script 907班後端 API 正常運作中！\n此網址為雲端資料庫接口，請將此網址複製並設定至 GitHub 的 index.html，即可啟用跨載具成績即時同步。")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * 支援外部獨立網頁或跨載具透過 POST 方式送出成績與管理指令
 */
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = e.parameter || {};
      }
    } else if (e.parameter) {
      payload = e.parameter;
    }

    if (payload.action === 'clearScores') {
      const res = clearAllScores(payload.password);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (payload.action === 'setRange') {
      const res = setActiveQuizRange(payload.password, payload.start, payload.end);
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const result = submitStudentScore(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 取得或自動建立 907 班成績試算表（利用 ScriptProperties 快取 Spreadsheet ID，保證所有載具 100% 同步開同一個檔案）
 */
function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      const ss = SpreadsheetApp.openById(savedId);
      if (ss) return ss;
    } catch (e) {
      Logger.log('無法藉由 ID 打開試算表：' + e.toString());
    }
  }

  const sheetName = '頭前國中_907班_英文單字小考成績登記表';
  let ss = null;

  try {
    const files = DriveApp.getFilesByName(sheetName);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      // 檢查舊檔名
      const oldFiles = DriveApp.getFilesByName('頭前國中_英文單字小考成績登記表(9/12~10/5)');
      if (oldFiles.hasNext()) {
        ss = SpreadsheetApp.open(oldFiles.next());
      }
    }
  } catch (driveErr) {
    Logger.log('DriveApp 權限受限：' + driveErr.toString());
  }

  if (!ss) {
    ss = SpreadsheetApp.create(sheetName);
    const sheet = ss.getActiveSheet();
    sheet.appendRow(['交卷時間', '班級', '座號', '姓名', '測驗起訖日期', '測驗得分 (滿分100)', '答錯單字/片語清單']);
    const headerRange = sheet.getRange(1, 1, 1, 7);
    headerRange.setBackground('#1d4ed8')
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  // 儲存 ID 快取，使後續所有載具連線瞬間定位
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

/**
 * 學生交卷時，自動將成績寫入雲端硬碟的試算表
 */
function submitStudentScore(payload) {
  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getActiveSheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    
    sheet.appendRow([
      timestamp,
      payload.className || '907',
      parseInt(payload.seatNo, 10) || payload.seatNo || '',
      payload.name || '',
      payload.range || '9/14~10/5',
      payload.score !== undefined ? Number(payload.score) : 0,
      payload.wrongWords || '無 (全部答對滿分)'
    ]);

    // 交卷後自動將所有資料依「座號」(第3欄) 由小到大排序，同座號依時間排序
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      if (lastRow > 2) {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
        dataRange.sort([{ column: 3, ascending: true }, { column: 1, ascending: true }]);
      }
      
      // 找出每個座號的最高分，並在試算表得分欄(第6欄)標示紅色粗體
      const seatValues = sheet.getRange(2, 3, lastRow - 1, 4).getValues(); // 座號(col 3), 姓名(col 4), 起訖(col 5), 得分(col 6)
      const maxScoreBySeat = {};
      for (let i = 0; i < seatValues.length; i++) {
        const s = String(seatValues[i][0]);
        const sc = Number(seatValues[i][3]) || 0;
        if (maxScoreBySeat[s] === undefined || sc > maxScoreBySeat[s]) {
          maxScoreBySeat[s] = sc;
        }
      }
      for (let i = 0; i < seatValues.length; i++) {
        const s = String(seatValues[i][0]);
        const sc = Number(seatValues[i][3]) || 0;
        const scoreCell = sheet.getRange(i + 2, 6);
        if (sc === maxScoreBySeat[s]) {
          scoreCell.setFontColor('#dc2626').setFontWeight('bold');
        } else {
          scoreCell.setFontColor('#000000').setFontWeight('normal');
        }
      }
    }

    // 4. 同步更新「907班_每人最高成績登記表」，永久保留每位學生測驗最高分
    updateSummarySheet(ss, payload);

    return { success: true, sheetUrl: ss.getUrl() };
  } catch (err) {
    Logger.log('儲存成績失敗：' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 取得所有學生的測驗時間、起訖日期與成績紀錄（依座號由小到大排序）
 * 任何同學在不同手機、平板、電腦皆可呼叫此函式查看全班成績
 */
function getAllScores() {
  try {
    const ss = getOrCreateSpreadsheet();
    if (!ss) return [];

    const sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const lastCol = Math.max(7, sheet.getLastColumn());
    const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // 動態偵測各欄位索引，相容所有歷史試算表格式
    let colTime = 0, colClass = 1, colSeat = 2, colName = 3, colRange = 4, colScore = 5, colWrong = 6;
    headerRow.forEach((h, idx) => {
      const hStr = String(h || '');
      if (hStr.includes('時間')) colTime = idx;
      else if (hStr.includes('班級')) colClass = idx;
      else if (hStr.includes('座號')) colSeat = idx;
      else if (hStr.includes('姓名')) colName = idx;
      else if (hStr.includes('起訖') || hStr.includes('範圍') || hStr.includes('日期')) colRange = idx;
      else if (hStr.includes('得分') || hStr.includes('成績') || hStr.includes('分數')) colScore = idx;
      else if (hStr.includes('錯') || hStr.includes('清單')) colWrong = idx;
    });

    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const list = values.map(row => {
      let timeStr = '';
      if (row[colTime] instanceof Date) {
        timeStr = Utilities.formatDate(row[colTime], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      } else {
        timeStr = String(row[colTime] || '');
      }
      return {
        time: timeStr,
        className: String(row[colClass] || '907'),
        seatNo: parseInt(row[colSeat], 10) || row[colSeat],
        name: String(row[colName] || ''),
        range: String(row[colRange] || '9/14~10/5'),
        score: (row[colScore] !== '' && !isNaN(Number(row[colScore]))) ? Number(row[colScore]) : 0,
        wrongWords: String(row[colWrong] || '')
      };
    });

    // 依座號由小到大排序，同座號依時間排序
    list.sort((a, b) => {
      const seatA = parseInt(a.seatNo, 10) || 999;
      const seatB = parseInt(b.seatNo, 10) || 999;
      if (seatA !== seatB) return seatA - seatB;
      return a.time.localeCompare(b.time);
    });

    return list;
  } catch (err) {
    Logger.log('讀取成績失敗：' + err.toString());
    return [];
  }
}

/**
 * 清除試算表中的所有成績紀錄（需驗證密碼 123456）
 */
function clearAllScores(password) {
  if (String(password).trim() !== '123456') {
    return { success: false, error: '密碼錯誤！無法清除成績紀錄。' };
  }

  try {
    const ss = getOrCreateSpreadsheet();
    if (ss) {
      // 1. 清空明細工作表歷程紀錄
      const sheet = ss.getActiveSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }

      // 2. 重置 907 班每人最高成績登記表
      const summarySheet = ss.getSheetByName('907班_每人最高成績登記表');
      if (summarySheet) {
        initSummarySheetData(summarySheet);
      }
    }
    return { success: true, message: '907 班全班成績已全數清空！' };
  } catch (err) {
    Logger.log('清除成績失敗：' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 取得老師設定的目前測驗起訖日期範圍（預設 9/14 ~ 10/5）
 */
function getActiveQuizRange() {
  try {
    const props = PropertiesService.getScriptProperties();
    const start = props.getProperty('QUIZ_START_DATE') || '9/14';
    const end = props.getProperty('QUIZ_END_DATE') || '10/5';
    return { success: true, start: start, end: end };
  } catch (err) {
    return { success: false, start: '9/14', end: '10/5' };
  }
}

/**
 * 老師設定測驗起訖日期範圍（僅限 38 號 劉真妮 老師，需驗證密碼 123456）
 */
function setActiveQuizRange(password, startDate, endDate) {
  if (String(password).trim() !== '123456') {
    return { success: false, error: '密碼錯誤！僅限 38 號 劉真妮 老師有權限更改測驗日期範圍。' };
  }
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('QUIZ_START_DATE', String(startDate));
    props.setProperty('QUIZ_END_DATE', String(endDate));
    props.setProperty('QUIZ_TEACHER', '38號 劉真妮');
    return { success: true, start: startDate, end: endDate, teacher: '38號 劉真妮' };
  } catch (err) {
    Logger.log('儲存測驗範圍失敗：' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 907 班官方名冊（01 ~ 31 號 及 38 號 劉真妮 老師）
 */
const ROSTER_907_DATA = [
  { seat: 1, name: "陳品澄" },
  { seat: 2, name: "潘宥任" },
  { seat: 3, name: "林秉樂" },
  { seat: 4, name: "麥藝霖" },
  { seat: 5, name: "郭宸良" },
  { seat: 6, name: "劉仲軒" },
  { seat: 7, name: "吳孟駿" },
  { seat: 8, name: "吳承翰" },
  { seat: 9, name: "楊家宥" },
  { seat: 10, name: "王陽晟" },
  { seat: 11, name: "彭貝宇" },
  { seat: 12, name: "宋昱翰" },
  { seat: 13, name: "劉邑璽" },
  { seat: 14, name: "周莛恩" },
  { seat: 15, name: "蘇子睿" },
  { seat: 16, name: "丘天玥" },
  { seat: 17, name: "林盈妡" },
  { seat: 18, name: "詹子晴" },
  { seat: 19, name: "謝沛琪" },
  { seat: 20, name: "曾沛繡" },
  { seat: 21, name: "賴律穎" },
  { seat: 22, name: "黃映綸" },
  { seat: 23, name: "賴鈺涵" },
  { seat: 24, name: "黃若涵" },
  { seat: 25, name: "吳采容" },
  { seat: 26, name: "邱筠雲" },
  { seat: 27, name: "柯亮瑀" },
  { seat: 28, name: "簡舒涵" },
  { seat: 29, name: "陳言沁" },
  { seat: 30, name: "馮羽婕" },
  { seat: 31, name: "潘彥蓁" },
  { seat: 38, name: "劉真妮" }
];

/**
 * 取得或初始化「907班_每人最高成績登記表」
 */
function getOrCreateSummarySheet(ss) {
  const sheetName = '907班_每人最高成績登記表';
  let summarySheet = ss.getSheetByName(sheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(sheetName);
    initSummarySheetData(summarySheet);
  }
  return summarySheet;
}

/**
 * 初始化最高成績登記表欄位與 907 班名冊
 */
function initSummarySheetData(sheet) {
  sheet.clear();
  sheet.appendRow(['班級', '座號', '姓名', '🏆 最高登記得分', '累積測驗次數', '最近一次得分', '最新測驗時間', '最新測驗起訖']);
  const header = sheet.getRange(1, 1, 1, 8);
  header.setBackground('#1d4ed8')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

  const rows = ROSTER_907_DATA.map(r => [
    '907',
    r.seat,
    r.name,
    '-',
    0,
    '-',
    '-',
    '-'
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
    sheet.getRange(2, 1, rows.length, 2).setHorizontalAlignment('center');
    sheet.getRange(2, 4, rows.length, 3).setHorizontalAlignment('center');
  }
  sheet.setFrozenRows(1);
}

/**
 * 更新「907班_每人最高成績登記表」
 * 自動比較歷次成績，永久保留學生最高得分，並以紅色粗體顯示
 */
function updateSummarySheet(ss, payload) {
  try {
    const summarySheet = getOrCreateSummarySheet(ss);
    const lastRow = summarySheet.getLastRow();
    if (lastRow < 2) {
      initSummarySheetData(summarySheet);
    }
    
    const seatNo = parseInt(payload.seatNo, 10);
    const score = Number(payload.score) || 0;
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const rangeStr = payload.range || '';
    
    const data = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, 8).getValues();
    let targetRowIndex = -1;
    
    for (let i = 0; i < data.length; i++) {
      if (parseInt(data[i][1], 10) === seatNo) {
        targetRowIndex = i + 2; // 試算表列號從 1 算起，第 1 列為標題
        const currentMax = Number(data[i][3]);
        const testCount = (Number(data[i][4]) || 0) + 1;
        
        let newMax = score;
        if (!isNaN(currentMax) && currentMax > score) {
          newMax = currentMax; // 維持歷史最高分
        }
        
        // 更新該列
        summarySheet.getRange(targetRowIndex, 4).setValue(newMax);
        summarySheet.getRange(targetRowIndex, 5).setValue(testCount);
        summarySheet.getRange(targetRowIndex, 6).setValue(score);
        summarySheet.getRange(targetRowIndex, 7).setValue(timestamp);
        summarySheet.getRange(targetRowIndex, 8).setValue(rangeStr);
        
        // 最高分以紅色粗體醒目標記
        summarySheet.getRange(targetRowIndex, 4)
          .setFontColor('#dc2626')
          .setFontWeight('bold');
        break;
      }
    }
    
    // 若在名冊中未找到（例如非 01~31 或 38），則自動追加
    if (targetRowIndex === -1) {
      summarySheet.appendRow([
        payload.className || '907',
        seatNo || payload.seatNo,
        payload.name || '',
        score,
        1,
        score,
        timestamp,
        rangeStr
      ]);
      const newRow = summarySheet.getLastRow();
      summarySheet.getRange(newRow, 4)
        .setFontColor('#dc2626')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  } catch (err) {
    Logger.log('更新最高成績表失敗：' + err.toString());
  }
}


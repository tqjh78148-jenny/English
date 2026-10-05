/**
 * 🏫 頭前國中 英文單字小考 — 雲端隨機抽考系統 (Google Apps Script 後端)
 *
 * 特色：
 * 1. 學生點開 Google 官方網址，每次進去「隨機重抽 10 題」（9/12～10/5 共 112 題庫）。
 * 2. 題目附帶 🔊 真人發音、即時計分（滿分 100 分）、錯題複習。
 * 3. 學生交卷後，系統自動將「班級、座號、姓名、得分、測驗時間、錯題」記錄到雲端試算表，老師直接登記成績！
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('頭前國中 英文聯絡簿單字小考 (9/12~10/5)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 支援外部獨立網頁透過 POST 方式送出成績
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const result = submitStudentScore(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 學生交卷時，自動將成績寫入雲端硬碟的試算表
 */
function submitStudentScore(payload) {
  try {
    const sheetName = '頭前國中_907班_英文單字小考成績登記表';
    let ss;
    const files = DriveApp.getFilesByName(sheetName);
    
    // 若試算表已存在則打開，若不存在則自動新建
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
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

    const sheet = ss.getActiveSheet();
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    
    sheet.appendRow([
      timestamp,
      payload.className || '907',
      parseInt(payload.seatNo, 10) || payload.seatNo || '',
      payload.name || '',
      payload.range || '9/14~10/5',
      payload.score !== undefined ? payload.score : 0,
      payload.wrongWords || '無 (全部答對滿分)'
    ]);

    // 交卷後自動將所有資料依「座號」(第3欄) 由小到大排序，同座號依時間排序
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      if (lastRow > 2) {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
        dataRange.sort([{ column: 3, ascending: true }, { column: 1, ascending: true }]);
      }
      
      // 找出每個座號的最高分，並在試算表得分欄(第6欄)標示紅色
      const seatValues = sheet.getRange(2, 3, lastRow - 1, 4).getValues(); // 取得 座號(col 3), 姓名(col 4), 起訖(col 5), 得分(col 6)
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

    return { success: true, sheetUrl: ss.getUrl() };
  } catch (err) {
    Logger.log('儲存成績失敗：' + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 取得所有學生的測驗時間、起訖日期與成績紀錄（依座號由小到大排序）
 */
function getAllScores() {
  try {
    const sheetNames = ['頭前國中_907班_英文單字小考成績登記表', '頭前國中_英文單字小考成績登記表(9/12~10/5)'];
    let ss = null;
    for (let i = 0; i < sheetNames.length; i++) {
      const files = DriveApp.getFilesByName(sheetNames[i]);
      if (files.hasNext()) {
        ss = SpreadsheetApp.open(files.next());
        break;
      }
    }
    if (!ss) return [];

    const sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const headerRow = sheet.getRange(1, 1, 1, 7).getValues()[0];
    const isCol5Range = String(headerRow[4] || '').includes('起訖') || String(headerRow[4] || '').includes('範圍');

    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const list = values.map(row => {
      let timeStr = '';
      if (row[0] instanceof Date) {
        timeStr = Utilities.formatDate(row[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
      } else {
        timeStr = String(row[0] || '');
      }
      return {
        time: timeStr,
        className: String(row[1] || '907'),
        seatNo: parseInt(row[2], 10) || row[2],
        name: String(row[3] || ''),
        range: isCol5Range ? String(row[4] || '') : String(row[5] || ''),
        score: isCol5Range ? row[5] : row[4],
        wrongWords: String(row[6] || '')
      };
    });

    // 依座號由小到大排序
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
    const sheetNames = ['頭前國中_907班_英文單字小考成績登記表', '頭前國中_英文單字小考成績登記表(9/12~10/5)'];
    sheetNames.forEach(sheetName => {
      const files = DriveApp.getFilesByName(sheetName);
      while (files.hasNext()) {
        const ss = SpreadsheetApp.open(files.next());
        const sheet = ss.getActiveSheet();
        const lastRow = sheet.getLastRow();
        // 保留第一行表頭，清除資料列
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
        }
      }
    });

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



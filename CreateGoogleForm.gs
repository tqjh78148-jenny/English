/**
 * 🏫 頭前國中 英文小考：Google 表單測驗卷 ＋ 自動成績登記試算表
 * 
 * 範圍：9月12日 ～ 10月5日（今日）聯絡簿進度
 * 滿分：100 分（抽考 10 題，每題 10 分）
 * 功能：學生點連結作答，交卷後「成績自動寫入 Google 試算表（Excel）」，老師直接登記成績！
 *
 * 【操作步驟】：
 * 1. 打開學校 Google 雲端硬碟 ➔ 點左上角「新增」➔「更多」➔「Google Apps Script」。
 * 2. 把本檔案全部內容複製貼上。
 * 3. 點選上方「執行」按鈕（第一次會跳出安全性授權，點選「進階」➔「前往專案(安全)」允許即可）。
 * 4. 執行完畢後，回到 Google 雲端硬碟，會自動多出兩個檔案：
 *    - 📝【表單】頭前國中 英文聯絡簿單字小考 (9/12~10/5) ➔ 傳給學生的測驗連結
 *    - 📊【試算表】頭前國中 英文單字小考成績登記表 (9/12~10/5) ➔ 老師看成績的 Excel 表
 */

function createEnglishQuizWithGradeSheet() {
  // 1. 建立測驗表單
  const formTitle = '頭前國中 英文聯絡簿單字小考 (9/12~10/5)';
  const form = FormApp.create(formTitle);
  form.setDescription('115學年度 聯絡簿單字片語小考\n範圍：9月12日 ～ 10月5日\n總分：100 分（共 10 題，每題 10 分）\n請同學填寫班級、座號與姓名後認真作答！')
      .setIsQuiz(true)
      .setAllowResponseEdits(false)
      .setLimitOneResponsePerUser(false);

  // 2. 自動在雲端硬碟建立並綁定「成績登記試算表」
  const sheetTitle = '【成績登記表】頭前國中 英文單字小考 (9/12~10/5)';
  const ss = SpreadsheetApp.create(sheetTitle);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 3. 學生身分題目
  form.addTextItem().setTitle('1. 班級 (例: 115)').setRequired(true);
  form.addTextItem().setTitle('2. 座號 (例: 01)').setRequired(true);
  form.addTextItem().setTitle('3. 姓名').setRequired(true);

  // 4. 嚴格限定 9/14 ～ 10/5 (今日) 題庫 (共 112 個單字與片語)
  const pool = [
    // 9/14
    { date: "9/14", word: "alive", meaning: "活著的", type: "單字" },
    { date: "9/14", word: "allow", meaning: "允許", type: "單字" },
    { date: "9/14", word: "almost", meaning: "幾乎", type: "單字" },
    { date: "9/14", word: "alone", meaning: "單獨地", type: "單字" },
    { date: "9/14", word: "along", meaning: "沿著", type: "單字" },
    { date: "9/14", word: "all at once", meaning: "突然地", type: "片語" },
    { date: "9/14", word: "all day long", meaning: "一整天", type: "片語" },
    { date: "9/14", word: "all in all", meaning: "總而言之", type: "片語" },

    // 9/15
    { date: "9/15", word: "already", meaning: "已經", type: "單字" },
    { date: "9/15", word: "also", meaning: "也", type: "單字" },
    { date: "9/15", word: "although", meaning: "雖然", type: "單字" },
    { date: "9/15", word: "always", meaning: "總是", type: "單字" },
    { date: "9/15", word: "amazing", meaning: "令人驚奇的", type: "單字" },
    { date: "9/15", word: "all kinds of", meaning: "各種各樣的", type: "片語" },
    { date: "9/15", word: "all of a sudden", meaning: "突然", type: "片語" },
    { date: "9/15", word: "all over", meaning: "到處；遍及", type: "片語" },

    // 9/16
    { date: "9/16", word: "amber", meaning: "琥珀色", type: "單字" },
    { date: "9/16", word: "ambition", meaning: "野心；抱負", type: "單字" },
    { date: "9/16", word: "ambulance", meaning: "救護車", type: "單字" },
    { date: "9/16", word: "among", meaning: "在……之中", type: "單字" },
    { date: "9/16", word: "amount", meaning: "數量", type: "單字" },
    { date: "9/16", word: "all the time", meaning: "總是", type: "片語" },
    { date: "9/16", word: "all the way", meaning: "一路上", type: "片語" },
    { date: "9/16", word: "along with", meaning: "與……一起", type: "片語" },

    // 9/17
    { date: "9/17", word: "ancient", meaning: "古老的", type: "單字" },
    { date: "9/17", word: "angry", meaning: "生氣的", type: "單字" },
    { date: "9/17", word: "animal", meaning: "動物", type: "單字" },
    { date: "9/17", word: "ankle", meaning: "腳踝", type: "單字" },
    { date: "9/17", word: "another", meaning: "另一個", type: "單字" },
    { date: "9/17", word: "answer for", meaning: "對……負責", type: "片語" },
    { date: "9/17", word: "apart from", meaning: "除了……之外", type: "片語" },
    { date: "9/17", word: "apologize to", meaning: "向……道歉", type: "片語" },

    // 9/18
    { date: "9/18", word: "answer", meaning: "回答；答案", type: "單字" },
    { date: "9/18", word: "ant", meaning: "螞蟻", type: "單字" },
    { date: "9/18", word: "anxious", meaning: "焦慮的", type: "單字" },
    { date: "9/18", word: "anybody", meaning: "任何人", type: "單字" },
    { date: "9/18", word: "anymore", meaning: "再也（不）", type: "單字" },
    { date: "9/18", word: "appeal to", meaning: "吸引；呼籲", type: "片語" },
    { date: "9/18", word: "apply for", meaning: "申請", type: "片語" },
    { date: "9/18", word: "apply to", meaning: "適用於", type: "片語" },

    // 9/21
    { date: "9/21", word: "anyone", meaning: "任何人", type: "單字" },
    { date: "9/21", word: "anything", meaning: "任何事", type: "單字" },
    { date: "9/21", word: "anyway", meaning: "無論如何", type: "單字" },
    { date: "9/21", word: "anywhere", meaning: "任何地方", type: "單字" },
    { date: "9/21", word: "apartment", meaning: "公寓", type: "單字" },
    { date: "9/21", word: "approve of", meaning: "贊成", type: "片語" },
    { date: "9/21", word: "argue with", meaning: "與……爭論", type: "片語" },
    { date: "9/21", word: "arm in arm", meaning: "手臂挽著手臂", type: "片語" },

    // 9/22
    { date: "9/22", word: "apologize", meaning: "道歉", type: "單字" },
    { date: "9/22", word: "appear", meaning: "出現；似乎", type: "單字" },
    { date: "9/22", word: "apple", meaning: "蘋果", type: "單字" },
    { date: "9/22", word: "apply", meaning: "申請；應用", type: "單字" },
    { date: "9/22", word: "appreciate", meaning: "感激；欣賞", type: "單字" },
    { date: "9/22", word: "around the corner", meaning: "即將來臨；在拐角處", type: "片語" },
    { date: "9/22", word: "arrive at", meaning: "到達（小地點）", type: "片語" },
    { date: "9/22", word: "arrive in", meaning: "到達（大城市/國家）", type: "片語" },

    // 9/23
    { date: "9/23", word: "area", meaning: "區域", type: "單字" },
    { date: "9/23", word: "argue", meaning: "爭論", type: "單字" },
    { date: "9/23", word: "arm", meaning: "手臂", type: "單字" },
    { date: "9/23", word: "army", meaning: "軍隊", type: "單字" },
    { date: "9/23", word: "around", meaning: "周圍；大約", type: "單字" },
    { date: "9/23", word: "as a matter of fact", meaning: "事實上", type: "片語" },
    { date: "9/23", word: "as a result", meaning: "結果；因此", type: "片語" },
    { date: "9/23", word: "as far as", meaning: "就……而言", type: "片語" },

    // 9/24
    { date: "9/24", word: "arrange", meaning: "安排", type: "單字" },
    { date: "9/24", word: "arrest", meaning: "逮捕", type: "單字" },
    { date: "9/24", word: "arrive", meaning: "到達", type: "單字" },
    { date: "9/24", word: "art", meaning: "藝術", type: "單字" },
    { date: "9/24", word: "article", meaning: "文章", type: "單字" },
    { date: "9/24", word: "as for", meaning: "至於", type: "片語" },
    { date: "9/24", word: "as if", meaning: "彷彿；好像", type: "片語" },
    { date: "9/24", word: "as long as", meaning: "只要", type: "片語" },

    // 9/29
    { date: "9/29", word: "artist", meaning: "藝術家", type: "單字" },
    { date: "9/29", word: "as", meaning: "作為；當……時", type: "單字" },
    { date: "9/29", word: "ashamed", meaning: "羞愧的", type: "單字" },
    { date: "9/29", word: "ask", meaning: "詢問；要求", type: "單字" },
    { date: "9/29", word: "asleep", meaning: "睡著的", type: "單字" },
    { date: "9/29", word: "as soon as", meaning: "一……就……", type: "片語" },
    { date: "9/29", word: "as well", meaning: "也", type: "片語" },
    { date: "9/29", word: "as well as", meaning: "以及；和", type: "片語" },

    // 9/30
    { date: "9/30", word: "assistant", meaning: "助手", type: "單字" },
    { date: "9/30", word: "assume", meaning: "假設；認為", type: "單字" },
    { date: "9/30", word: "attack", meaning: "攻擊", type: "單字" },
    { date: "9/30", word: "attend", meaning: "參加", type: "單字" },
    { date: "9/30", word: "attention", meaning: "注意力", type: "單字" },
    { date: "9/30", word: "ask for", meaning: "要求；索取", type: "片語" },
    { date: "9/30", word: "at a loss", meaning: "茫然不知所措", type: "片語" },
    { date: "9/30", word: "at all", meaning: "根本；一點也", type: "片語" },

    // 10/1
    { date: "10/1", word: "attitude", meaning: "態度", type: "單字" },
    { date: "10/1", word: "attract", meaning: "吸引", type: "單字" },
    { date: "10/1", word: "audience", meaning: "觀眾", type: "單字" },
    { date: "10/1", word: "aunt", meaning: "姑姑；阿姨", type: "單字" },
    { date: "10/1", word: "author", meaning: "作者", type: "單字" },
    { date: "10/1", word: "at all costs", meaning: "不惜任何代價", type: "片語" },
    { date: "10/1", word: "at best", meaning: "充其量", type: "片語" },
    { date: "10/1", word: "at first", meaning: "起初", type: "片語" },

    // 10/2
    { date: "10/2", word: "autumn", meaning: "秋天", type: "單字" },
    { date: "10/2", word: "available", meaning: "可得到的；有空的", type: "單字" },
    { date: "10/2", word: "average", meaning: "平均的", type: "單字" },
    { date: "10/2", word: "avoid", meaning: "避免", type: "單字" },
    { date: "10/2", word: "award", meaning: "獎項；頒發", type: "單字" },
    { date: "10/2", word: "at hand", meaning: "在手邊；即將到來", type: "片語" },
    { date: "10/2", word: "at home", meaning: "在家；自在", type: "片語" },
    { date: "10/2", word: "at last", meaning: "終於", type: "片語" },

    // 10/5 (今日進度)
    { date: "10/5", word: "baby", meaning: "嬰兒", type: "單字" },
    { date: "10/5", word: "baby-sit", meaning: "托兒照顧", type: "單字" },
    { date: "10/5", word: "back", meaning: "背部；返回", type: "單字" },
    { date: "10/5", word: "background", meaning: "背景", type: "單字" },
    { date: "10/5", word: "backward", meaning: "向後地", type: "單字" },
    { date: "10/5", word: "at least", meaning: "至少", type: "片語" },
    { date: "10/5", word: "at most", meaning: "至多", type: "片語" },
    { date: "10/5", word: "at night", meaning: "在晚上", type: "片語" }
  ];

  // 隨機抽 10 題（含單字與重點片語）
  const shuffled = pool.sort(() => 0.5 - Math.random()).slice(0, 10);

  shuffled.forEach((item, idx) => {
    // 隨機挑 3 個干擾選項
    const others = pool.filter(p => p.word !== item.word)
                       .sort(() => 0.5 - Math.random())
                       .slice(0, 3);
    const options = [item.meaning, ...others.map(o => o.meaning)].sort(() => 0.5 - Math.random());

    const mcItem = form.addMultipleChoiceItem();
    mcItem.setTitle(`第 ${idx + 1} 題 [${item.date} ${item.type}]：請選出「${item.word}」的中文意思`)
          .setPoints(10)
          .setRequired(true);

    const choices = options.map(opt => {
      const isCorrect = (opt === item.meaning);
      return mcItem.createChoice(opt, isCorrect);
    });
    mcItem.setChoices(choices);
  });

  Logger.log('====================================');
  Logger.log('✅ 表單與成績登記表建立成功！');
  Logger.log('🔗 學生測驗網址：' + form.getPublishedUrl());
  Logger.log('📊 老師成績試算表：' + ss.getUrl());
  Logger.log('✏️ 老師編輯表單網址：' + form.getEditUrl());
  Logger.log('====================================');
}

/**
 * 我的挫折思考小測驗 — Google Apps Script 後端（網頁放在 GitHub Pages）
 * 請綁定在要存資料的 Google 試算表上（擴充功能 → Apps Script）。
 * 部署：網頁應用程式／執行身分：我／存取：所有人
 * 教師密碼請設在「專案設定 → 指令碼屬性」：屬性 TEACHER_PASSWORD
 *  （設一次就好，之後重新貼上程式也不用再改；密碼不會出現在 GitHub）
 */
var TEACHER_PASSWORD = 'change-me'; // 第 8 行：沒設指令碼屬性時才會用到

var SHEET_NAME = '前測資料';
var QUESTION_SHEET_NAME = '題目對照';
var TIMEZONE = 'Asia/Taipei';
var MAX_LEN = 3000;

var HEADERS = [
  '提交時間', '班級', '座號或研究代碼', '情境代號', '情境標題',
  '選擇原因代號', '選擇原因', '其他原因說明',
  'Q1_理解主角感受', 'Q2_可能的原因', 'Q3_原因以後會不會改變', 'Q4_可改變與不可控制',
  'Q5_重新看待這件事', 'Q6_接下來的做法',
  '開始作答時間', '作答秒數', '裝置', '提交ID', '版本', '備註',
  'C3_原因以後會', 'C4_主要能否控制', 'C5_與主角想法比較', 'C6_主角之後比較適合'
];

/* 第3–6題開放題寫完後的補充選擇題（文字須與網頁一致） */
var CHOICES = {
  c3: { label: 'C3_原因以後會', prompt: '你覺得這次的原因，以後會？', options: ['大概會一直存在', '可能會隨情況改變', '不確定'] },
  c4: { label: 'C4_主要能否控制', prompt: '整體來說，這件事主要是？', options: ['主角自己可以控制的', '主角不太能控制的', '兩者都有', '不確定'] },
  c5: { label: 'C5_與主角想法比較', prompt: '你對這次經驗的看法，和主角原本的想法比較起來？', options: ['差不多相同', '部分相同', '不太一樣', '不確定'] },
  c6: { label: 'C6_主角之後比較適合', prompt: '你覺得主角之後比較適合？', options: ['用原本的方式再試一次', '換個方法再試', '先暫停，之後再決定', '不再嘗試這件事', '不確定'] }
};

var SCENARIOS = {
  A: '努力了，成績卻不如預期',
  B: '比賽輸了，覺得自己拖累大家',
  C: '花了很多心思，作品卻沒有入選',
  D: '朋友好像漸漸疏遠了'
};

var REASONS = {
  '1': '我曾經有過類似的經驗。',
  '2': '我很能理解故事中人物的感受。',
  '3': '我覺得這種情況在生活中很常見。',
  '4': '我只是對這個故事比較有興趣。',
  '5': '其他原因'
};

/* ---------- 進入點 ---------- */
// 直接打開 /exec 網址會看到 {"ok":true,...}，這是正常的（網頁在 GitHub Pages）
function doGet() {
  return json_({ ok: true, service: 'frustration-pretest' });
}

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'save';
    if (action === 'save') {
      result = submitResponse(body.data || body);
    } else if (action === 'list') {
      requirePw_(body.password);
      result = { ok: true, rows: listRows_() };
    } else if (action === 'progress') {
      requirePw_(body.password);
      result = { ok: true, classes: progress_() };
    } else {
      result = { ok: false, error: '未知的動作' };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  return json_(result);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* 密碼由後端驗證；預設 change-me 一律視為錯誤 */
function requirePw_(pw) {
  var real = PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD') || TEACHER_PASSWORD;
  if (!real || real === 'change-me' || String(pw || '') !== String(real)) {
    throw new Error('密碼錯誤');
  }
}

/* 教師用：讀出全部作答（只有密碼正確才會執行） */
var ROW_KEYS = ['time', 'className', 'code', 'scenario', 'scenarioTitle', 'reasonCode', 'reason', 'reasonOther',
  'q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'startedAt', 'seconds', 'device', 'id', 'version', 'note',
  'c3', 'c4', 'c5', 'c6'];

function listRows_() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  return values.filter(function (r) { return r[17]; }).map(function (r) {
    var o = {};
    ROW_KEYS.forEach(function (k, i) { o[k] = r[i]; });
    return o;
  });
}

/* 大屏用：每班已完成的座號（不含任何作答內容） */
function progress_() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var out = {};
  if (last < 2) return out;
  sheet.getRange(2, 2, last - 1, 2).getDisplayValues().forEach(function (r) {
    var c = String(r[0]), code = String(r[1]);
    if (!c) return;
    out[c] = out[c] || [];
    if (out[c].indexOf(code) < 0) out[c].push(code);
  });
  return out;
}

/* ---------- 寫入資料（網頁透過 google.script.run 呼叫） ---------- */
function submitResponse(data) {
  var clean = validate_(data);           // 檢查不過會丟出錯誤
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                   // 多人同時送出時排隊寫入
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    var idCol = HEADERS.indexOf('提交ID') + 1;

    // 防重複：同一個提交ID已經存在就不再寫入
    if (lastRow > 1) {
      var found = sheet.getRange(2, idCol, lastRow - 1, 1)
        .createTextFinder(clean.submissionId).matchEntireCell(true).findNext();
      if (found) return { ok: true, duplicate: true };
    }

    // 同班級＋同代碼已有資料時，只做標記，不擋下（交給研究者判斷）
    var note = '';
    if (lastRow > 1) {
      var keys = sheet.getRange(2, 2, lastRow - 1, 2).getValues();
      var n = 0;
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === clean.className && String(keys[i][1]) === clean.studentCode) n++;
      }
      if (n > 0) note = '同班級與代碼已有 ' + n + ' 筆資料';
    }

    var now = new Date();
    var started = clean.startedAt ? new Date(clean.startedAt) : null;
    var seconds = (started && !isNaN(started)) ? Math.round((now - started) / 1000) : '';

    var row = [
      fmt_(now),
      clean.className,
      clean.studentCode,
      clean.scenario,
      SCENARIOS[clean.scenario],
      clean.reasonCode,
      REASONS[clean.reasonCode],
      clean.reasonOther,
      clean.q1, clean.q2, clean.q3, clean.q4, clean.q5, clean.q6,
      (started && !isNaN(started)) ? fmt_(started) : '',
      seconds,
      clean.device,
      clean.submissionId,
      clean.version,
      note,
      clean.c3, clean.c4, clean.c5, clean.c6
    ].map(safeCell_);

    sheet.appendRow(row);
    SpreadsheetApp.flush();
    return { ok: true, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 伺服器端檢查（不信任前端） ---------- */
function validate_(d) {
  if (!d || typeof d !== 'object') throw new Error('沒有收到資料');
  var s = function (v) { return (v === undefined || v === null) ? '' : String(v); };
  var out = {
    submissionId: s(d.submissionId).trim().slice(0, 80),
    version: s(d.version).slice(0, 40),
    className: s(d.className).trim().slice(0, 20),
    studentCode: s(d.studentCode).trim().slice(0, 20),
    scenario: s(d.scenario).trim().toUpperCase(),
    reasonCode: s(d.reasonCode).trim(),
    reasonOther: s(d.reasonOther).slice(0, 300),
    startedAt: s(d.startedAt).slice(0, 40),
    device: s(d.device).slice(0, 10)
  };
  // 六題原始答案：完整保留（只限制最長長度）
  for (var i = 1; i <= 6; i++) out['q' + i] = s(d['q' + i]).slice(0, MAX_LEN);
  Object.keys(CHOICES).forEach(function (k) { out[k] = s(d[k]).trim(); });

  if (!out.submissionId) throw new Error('缺少提交ID');
  if (!out.className) throw new Error('缺少班級');
  if (!out.studentCode) throw new Error('缺少座號或研究代碼');
  if (!SCENARIOS[out.scenario]) throw new Error('情境代號不正確');
  if (!REASONS[out.reasonCode]) throw new Error('選擇原因不正確');
  if (out.reasonCode === '5' && !out.reasonOther.trim()) throw new Error('缺少其他原因說明');
  if (out.reasonCode !== '5') out.reasonOther = '';
  for (var j = 1; j <= 6; j++) {
    if (!out['q' + j].trim()) throw new Error('第 ' + j + ' 題沒有作答');
  }
  Object.keys(CHOICES).forEach(function (k) {
    if (CHOICES[k].options.indexOf(out[k]) < 0) throw new Error(CHOICES[k].label + ' 沒有選或選項不正確');
  });
  return out;
}

/* 避免學生文字以 = + - @ 開頭時被試算表當成公式 */
function safeCell_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function fmt_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setupSheet(); sheet = ss.getSheetByName(SHEET_NAME); }
  // 欄位有新增時（例如加了 C3–C6），自動補上標題列
  if (sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e4edf3').setWrap(true);
  }
  return sheet;
}

/* ---------- 第一次使用時執行一次：建立欄位與題目對照表 ---------- */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#e4edf3').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  sheet.setColumnWidths(1, 8, 120);
  sheet.setColumnWidths(9, 6, 320);
  sheet.getRange(2, 9, sheet.getMaxRows() - 1, 6).setWrap(true).setVerticalAlignment('top');
  sheet.setColumnWidths(21, 4, 160);
  sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@'); // 班級、代碼保持文字（例如 012 不會變成 12）

  var q = ss.getSheetByName(QUESTION_SHEET_NAME) || ss.insertSheet(QUESTION_SHEET_NAME);
  q.clear();
  var rows = [
    ['欄位', '題目名稱', '完整題目', '提示文字'],
    ['Q1_理解主角感受', '試著理解主角', '你認為故事中的主角現在可能有什麼感受？你從故事中的哪些地方這樣想？', '可以回到故事裡看看，是哪些句子或細節讓你這樣想。'],
    ['Q2_可能的原因', '事情為什麼會這樣？', '你認為這件事情發生的原因可能有哪些？除了故事中提到的情況，還有沒有其他可能？', '原因可能不只一個，想到幾個就寫幾個。'],
    ['Q3_原因以後會不會改變', '這些原因會一直存在嗎？', '你認為造成這次困難的原因，以後還會一直存在嗎？哪些可能改變？哪些可能不容易改變？為什麼？', '可以把想到的原因分開來想，一個一個說說看。'],
    ['Q4_可改變與不可控制', '哪些事情可以改變？', '在這個情境中，哪些事情是主角自己可以改變的？哪些事情不一定能由他控制？', '試著分別想想看，哪些事情可以靠自己調整，哪些事情可能受到其他因素影響。'],
    ['Q5_重新看待這件事', '重新看看這件事', '如果重新看待這件事，你會怎麼理解主角遇到的困難？你認為還可以有哪些不同的看法？為什麼？', '你的看法可以和主角一樣，也可以不一樣。'],
    ['Q6_接下來的做法', '接下來可以怎麼做？', '如果主角之後再次遇到類似的情況，你認為他可以怎麼做？也可以選擇不再嘗試，請說明你的理由。', '沒有一定要選擇哪一種做法，重要的是說明你的想法。'],
    ['', '', '', ''],
    ['情境代號', '情境標題', '人物', ''],
    ['A', SCENARIOS.A, '小凱', ''],
    ['B', SCENARIOS.B, '小宇', ''],
    ['C', SCENARIOS.C, '小晴', ''],
    ['D', SCENARIOS.D, '小安', ''],
    ['', '', '', ''],
    ['原因代號', '選擇原因', '', ''],
    ['1', REASONS['1'], '', ''], ['2', REASONS['2'], '', ''], ['3', REASONS['3'], '', ''],
    ['4', REASONS['4'], '', ''], ['5', '其他原因（文字見「其他原因說明」欄）', '', ''],
    ['', '', '', ''],
    ['補充選擇題欄位', '出現位置', '題目', '選項']
  ].concat(Object.keys(CHOICES).map(function (k) {
    return [CHOICES[k].label, '第' + k.slice(1) + '題開放題寫完後', CHOICES[k].prompt, CHOICES[k].options.join('／')];
  }));
  q.getRange(1, 1, rows.length, 4).setValues(rows).setWrap(true).setVerticalAlignment('top');
  [1, 9, 15, 22].forEach(function (r) { q.getRange(r, 1, 1, 4).setFontWeight('bold').setBackground('#e5efe7'); });
  q.setColumnWidth(1, 150); q.setColumnWidth(2, 220); q.setColumnWidth(3, 420); q.setColumnWidth(4, 360);
}

/* ---------- 測試用：在編輯器中執行，確認能寫入 ---------- */
function testSubmit() {
  var base = {
    submissionId: 'TEST-' + new Date().getTime(), version: 'test',
    className: 'TEST', studentCode: '00', reasonCode: '5', reasonOther: '測試用的其他原因',
    q1: '測試回答1', q2: '=這行以等號開頭', q3: '-以減號開頭', q4: '測試回答4',
    q5: '測試回答5\n第二行', q6: '測試回答6',
    c3: '不確定', c4: '兩者都有', c5: '部分相同', c6: '不再嘗試這件事',
    startedAt: new Date(Date.now() - 600000).toISOString(), device: '電腦'
  };
  ['A', 'B', 'C', 'D'].forEach(function (k) {
    var d = JSON.parse(JSON.stringify(base));
    d.scenario = k; d.submissionId = base.submissionId + '-' + k;
    Logger.log(k + ' → ' + JSON.stringify(submitResponse(d)));
  });
  // 重複送出同一個ID，應回傳 duplicate:true 且不會多一筆
  var again = JSON.parse(JSON.stringify(base)); again.scenario = 'A'; again.submissionId = base.submissionId + '-A';
  Logger.log('重複送出 → ' + JSON.stringify(submitResponse(again)));
  // 缺題目應該被擋下
  try {
    var bad = JSON.parse(JSON.stringify(base)); bad.scenario = 'B'; bad.q4 = '  '; bad.submissionId = 'TEST-BAD';
    submitResponse(bad);
    Logger.log('⚠ 缺題沒有被擋下');
  } catch (e) { Logger.log('缺題被擋下（正確）：' + e.message); }
}

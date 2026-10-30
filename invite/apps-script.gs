/**
 * 諾貝爾講座回函收件程式（貼到 Google 試算表 → 擴充功能 → Apps Script）
 * - 回覆寫入「回函」工作表；同一姓名再送一次，會更新原本那一列（以最新回覆為準）
 * - 網址加 ?action=stats 會回傳出席統計與名單
 * 欄位：A 最後回覆時間｜B 姓名｜C 單位｜D 參加講座｜E 參加晚宴｜F 回覆次數
 */
var SHEET_NAME = '回函';
var HEADER = ['最後回覆時間', '姓名', '單位', '參加諾貝爾得主講座', '參加諾貝爾得主歡迎晚宴', '回覆次數'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  sh.setFrozenRows(1);
  return sh;
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  var name = String(p.name || '').trim().slice(0, 60);
  var unit = String(p.unit || '').trim().slice(0, 80);
  var lecture = p.lecture === '參加' ? '參加' : '不克參加';
  var dinner = p.dinner === '參加' ? '參加' : '不克參加';
  if (!name) return json_({ ok: false, error: 'name required' });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = sheet_();
    var now = new Date();
    var last = sh.getLastRow();
    if (last > 1) {
      var names = sh.getRange(2, 2, last - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (String(names[i][0]).trim() === name) {
          var row = i + 2;
          var count = Number(sh.getRange(row, 6).getValue()) || 1;
          sh.getRange(row, 1, 1, 6).setValues([[now, name, unit, lecture, dinner, count + 1]]);
          return json_({ ok: true, updated: true });
        }
      }
    }
    sh.appendRow([now, name, unit, lecture, dinner, 1]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var sh = sheet_();
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
  rows = rows.filter(function (r) { return String(r[1]).trim() !== ''; });
  var s = { total: rows.length, lecture: 0, dinner: 0, both: 0, neither: 0,
            lectureNames: [], dinnerNames: [], declinedNames: [], updated: new Date().toISOString() };
  rows.forEach(function (r) {
    var name = String(r[1]).trim();
    var unit = String(r[2]).trim();
    var label = unit ? name + '（' + unit + '）' : name;
    var l = r[3] === '參加', d = r[4] === '參加';
    if (l) { s.lecture++; s.lectureNames.push(label); }
    if (d) { s.dinner++; s.dinnerNames.push(label); }
    if (l && d) s.both++;
    if (!l && !d) { s.neither++; s.declinedNames.push(label); }
  });
  return json_(s);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * يستقبل إجابات الاستبيان ويضيف كل إجابة صفًا في أول ورقة من الجدول.
 * الخطوات في README.md.
 */
var HEADERS = ['التاريخ', 'الاسم', 'القسم أو الإدارة', 'العملية المطلوب أتمتتها', 'العمليات الضرورية', 'عمليات أخرى', 'ملاحظات أو اقتراحات'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var d = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setRightToLeft(true);
    } else if (sheet.getRange(1, HEADERS.length).getValue() === '') {
      sheet.getRange(1, HEADERS.length).setValue(HEADERS[HEADERS.length - 1]).setFontWeight('bold');
    }
    sheet.appendRow([
      new Date(),
      clean(d.name),
      clean(d.dept),
      clean(d.challenges),
      clean(Array.isArray(d.processes) ? d.processes.join('، ') : ''),
      clean(d.more),
      clean(d.notes)
    ]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

/* يقصّ النص، ويمنع تفسيره كمعادلة في الجدول */
function clean(v) {
  v = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * مدارس جواثا الأهلية — سجل تقييم الإذاعة المدرسية (نسخة مستقلة)
 *
 * استخدم هذه النسخة إذا تعذّر فتح Apps Script من داخل جوجل شيت:
 *   1) افتح script.new وأنت داخل بحساب jawatha الصحيح.
 *   2) الصق هذا الملف كاملًا.
 *   3) Deploy  >  New deployment  >  Web app
 *      Execute as: Me        Who has access: Anyone
 *   4) انسخ الرابط المنتهي بـ /exec.
 *
 * الفرق الوحيد عن Code.gs أن هذه النسخة تفتح ملف الشيت برقمه،
 * فلا تحتاج أن يكون السكربت مربوطًا بالملف.
 */

/* ملف جوجل شيت الذي تُحفظ فيه التقييمات */
var SPREADSHEET_ID = '1TmQ_Ii5Sp5FkjWnxNJ4OAsDIzf7GoRzN2hEFwdpPxWA';

/* اكتب هنا نفس رمز دخول المشرف في الموقع — داخل Apps Script فقط، لا ترفعه إلى GitHub */
var ADMIN_KEY = 'ضع-رمز-المشرف-هنا';

var SHEET_NAME = 'التقييمات';

var HEADERS = [
  'وقت الإرسال','المعرف','المعلم','المدرسة','التاريخ','عنوان الإذاعة',
  '1. الالتزام بالوقت والتنظيم',
  '2. جودة المقدمة وجذب الانتباه',
  '3. تنوع الفقرات وتكاملها',
  '4. سلامة اللغة العربية والإلقاء',
  '5. جودة الصوت ووضوحه',
  '6. الإبداع والابتكار في التقديم',
  '7. القيمة التربوية والثقافية للمحتوى',
  '8. مشاركة الطلاب وتوزيع الأدوار',
  '9. استخدام الوسائل والتقنيات المناسبة',
  '10. التأثير في الجمهور وترك أثر إيجابي',
  'المجموع من 100','الملاحظات','الفصل المقيَّم'
];

function sheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  } else if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function score_(scores, i) {
  if (!scores || scores.length <= i) return 0;
  return Number(scores[i]) || 0;
}

function rowOf_(data) {
  var row = [new Date(), data.id || '', data.teacher || '', data.school || '',
             String(data.day || ''), data.topic || ''];
  for (var i = 0; i < 10; i++) row.push(score_(data.scores, i));
  row.push(Number(data.total) || 0);
  row.push(data.notes || '');
  row.push(data.klass || '');
  return row;
}

/** استقبال تقييم معلم: تقييم جديد يُضاف، وتقييم بنفس المعرف يُحدَّث */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var sh = sheet_();
    var at = -1;
    if (data.id && sh.getLastRow() > 1) {
      var ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) { at = i + 2; break; }
      }
    }
    if (data.action === 'delete') {
      if (data.key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });
      if (at > 0) sh.deleteRow(at);
      return json_({ ok: true, deleted: at > 0 });
    }
    var row = rowOf_(data);
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ ok: true, updated: at > 0 });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

/** قراءة كل التقييمات — للوحة النتائج فقط، ومحمية بالرمز */
function doGet(e) {
  var key = (e && e.parameter && e.parameter.key) || '';
  if (key !== ADMIN_KEY) return json_({ ok: false, error: 'unauthorized' });

  var sh = sheet_();
  if (sh.getLastRow() < 2) return json_({ ok: true, rows: [] });

  var values = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  var tz = Session.getScriptTimeZone();
  var rows = values.map(function (r) {
    var day = r[4];
    if (day instanceof Date) day = Utilities.formatDate(day, tz, 'yyyy-MM-dd');
    var at = r[0];
    if (at instanceof Date) at = at.toISOString();
    var scores = [];
    for (var i = 6; i <= 15; i++) scores.push(Number(r[i]) || 0);
    return {
      at: String(at), id: String(r[1]), teacher: String(r[2]), school: String(r[3]),
      day: String(day), topic: String(r[5]), scores: scores,
      total: Number(r[16]) || 0, notes: String(r[17] || ''), klass: String(r[18] || '')
    };
  });
  return json_({ ok: true, rows: rows });
}

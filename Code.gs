//***** Pitmaster — סיכום משתה פלור (אלונים) *****/
var SPREADSHEET_ID = '1bZeFjq47CVGF4MQOH2BWN4OBHVcowHHC6gdDyYBk3kI';
var SHEET_NAME     = 'סיכום משתה אלונים ';

var EMAILS = [
  'alonim.office@pitmaster.show',
  'rotem@pitmaster.show',
  'alonim@pitmaster.show',
  'office@pitmaster.show',
  'ido@pitmaster.show',
  // 'manager@example.com'
];

var ERROR_EMAILS = [
  'alonim.office@pitmaster.show',
];

var FILE_ID = '115aOjT_G96CWZyA4Kjz27o1R4S7VA-dc';

var LOG_SHEET_NAME = 'דוח הפעלות';
var MAIL_LOG_SHEET_NAME = 'לוג מיילים';

var SAVE_PDF_TO_DRIVE = false;
var PDF_FOLDER_ID     = '';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('סיכום משתה פלור — פיטמאסטר אלונים')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLogoDataUrl() {
  var blob = DriveApp.getFileById(FILE_ID).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function ensureHeaders_(sh, HEADERS) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    return;
  }
  var firstRow = sh.getRange(1,1,1,HEADERS.length).getValues()[0];
  var same = HEADERS.every(function(h,i){ return firstRow[i] === h; });
  if (!same) {
    sh.insertRows(1, 1);
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
  }
}

function ensureLogSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
  var headers = [
    'Timestamp','סטטוס','סיבה/שגיאה','כתובת עמוד','User-Agent',
    'תאריך','שעה','מנהל','פיט',
    'סה״כ נמענים','נשלח בהצלחה','כשלו','נמענים שנשלח','נמענים שנכשל'
  ];
  if (log.getLastRow() === 0) {
    log.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return log;
}

/* כתיבה ללוג מיילים לכל נמען */
function ensureMailLogSheet_(){
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(MAIL_LOG_SHEET_NAME) || ss.insertSheet(MAIL_LOG_SHEET_NAME);
  var headers = ['Timestamp','נמען','סטטוס','סיבה/שגיאה','תאריך','שעה','מנהל','פיט'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}

function logRun_(status, reason, meta, mailStat){
  try{
    var log = ensureLogSheet_();
    mailStat = mailStat || {total:0, ok:0, fail:0, sentList:[], failList:[]};
    log.appendRow([
      new Date(),
      status || '',
      reason || '',
      (meta && meta.pageUrl) || '',
      (meta && meta.ua) || '',
      (meta && meta.date) || '',
      (meta && meta.time) || '',
      (meta && meta.manager) || '',
      (meta && meta.pit) || '',
      mailStat.total,
      mailStat.ok,
      mailStat.fail,
      (mailStat.sentList || []).join(', '),
      (mailStat.failList || []).join(', ')
    ]);
  } catch(e){
    Logger.log('logRun_ failed: '+e);
  }
}

function logMailPerRecipient_(list, status, reason, p){
  try{
    var sh = ensureMailLogSheet_();
    var now = new Date();
    (list || []).forEach(function(addr){
      sh.appendRow([
        now,
        addr || '',
        status || '',
        reason || '',
        (p && p.date) || '',
        (p && p.time) || '',
        (p && p.manager) || '',
        (p && p.pit) || ''
      ]);
    });
  } catch(e){
    Logger.log('logMailPerRecipient_ failed: '+e);
  }
}

function _hashPayload_(p) {
  var toHash = {
    date: p && p.date, time: p && p.time, manager: p && p.manager, pit: p && p.pit,
    courses: p && p.courses, issuesText: p && p.issuesText, guestsText: p && p.guestsText,
    tablesText: p && p.tablesText, openedText: p && p.openedText, missingText: p && p.missingText,
    floorText: p && p.floorText, floorCustomerFeedbackText: p && p.floorCustomerFeedbackText,
    kitchenText: p && p.kitchenText, alcoholSoldText: p && p.alcoholSoldText,
    alcoholPouredText: p && p.alcoholPouredText, buildFaultsText: p && p.buildFaultsText,
    dishwashersConductText: p && p.dishwashersConductText, discountsText: p && p.discountsText,
    pitmasterSummaryText: p && p.pitmasterSummaryText, extraNotesText: p && p.extraNotesText
  };
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(toHash));
  return Utilities.base64Encode(bytes);
}

function validatePayload_(p){
  if (!p || typeof p !== 'object') {
    return { ok:false, reason: 'payload חסר' };
  }
  var missing = [];
  ['date','time','manager','pit'].forEach(function(k){
    if (!p[k]) missing.push(k);
  });
  var hasCourse = Array.isArray(p.courses) &&
                  p.courses.some(function(c){ return c && c.name && c.details; });
  if (!hasCourse) missing.push('courses');
  if (missing.length) return { ok:false, reason: 'חסרים שדות: ' + missing.join(', ') };
  return { ok:true };
}

function buildMailContent_(p, coursesText) {
  function nv(x){ return (x===undefined || x===null) ? '' : String(x); }
  function esc(s){ s = nv(s); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  var brandBg   = '#0B0F17', cardBg='#FFFFFF', accent='#D97706', heading='#111827', bodyColor='#374151', border='#E5E7EB';

  function block(title, content){
    return '<div style="margin:0 0 18px 0;padding-bottom:12px;border-bottom:1px solid '+border+';">'
         +   '<div style="font-size:20px;font-weight:800;color:'+heading+';margin-bottom:8px;font-family:Heebo,Arial,Helvetica,sans-serif;text-align:right;">'+ esc(title) +'</div>'
         +   '<div style="font-size:16px;line-height:1.65;color:'+bodyColor+';white-space:pre-wrap;font-family:Heebo,Arial,Helvetica,sans-serif;">'+ esc(content) +'</div>'
         + '</div>';
  }

  var text =
    'תאריך: ' + nv(p.date) + '  שעה: ' + nv(p.time) + '\n' +
    'שם המנהל: ' + nv(p.manager) + '\n' +
    'שם הפיט: ' + nv(p.pit) + '\n\n' +
    'פירוט מנות:\n' + nv(coursesText) + '\n\n' +
    'בעיות במהלך המשתה:\n' + nv(p.issuesText) + '\n\n' +
    'כמות סועדים:\n' + nv(p.guestsText) + '\n\n' +
    'מספרי שולחנות והמלצרים:\n' + nv(p.tablesText) + '\n\n' +
    'דברים שנפתחו:\n' + nv(p.openedText) + '\n\n' +
    'חוסרים:\n' + nv(p.missingText) + '\n\n' +
    'צוות פלור:\n' + nv(p.floorText) + '\n\n' +
    'חוות דעת לקוחות – צוות פלור:\n' + nv(p.floorCustomerFeedbackText) + '\n\n' +
    'צוות מטבח:\n' + nv(p.kitchenText) + '\n\n' +
    'אלכוהול שנמכר:\n' + nv(p.alcoholSoldText) + '\n\n' +
    'אלכוהול שנמזג:\n' + nv(p.alcoholPouredText) + '\n\n' +
    'תקלות בינוי:\n' + nv(p.buildFaultsText) + '\n\n' +
    'התנהלות שוטפי כלים:\n' + nv(p.dishwashersConductText) + '\n\n' +
    'הנחות/OTH:\n' + nv(p.discountsText) + '\n\n' +
    'סיכום הפיטמאסטר:\n' + nv(p.pitmasterSummaryText) + '\n\n' +
    'הערות נוספות:\n' + nv(p.extraNotesText);

  var html =
    '<div style="direction:rtl;background:'+brandBg+';padding:0;margin:0;font-family:Heebo,Arial,Helvetica,sans-serif;">'
    + '<div style="max-width:800px;margin:0 auto;padding:12px 16px;border-bottom:4px solid '+accent+';"></div>'
    + '<div style="max-width:800px;margin:16px auto 24px auto;background:'+cardBg+';padding:20px 18px;border-radius:14px;border:1px solid '+border+';box-shadow:0 8px 24px rgba(0,0,0,.08);">'
    +   '<div style="text-align:center;margin-bottom:12px;">'
    +     '<img src="cid:logoImage" alt="Pitmaster" style="max-height:120px;border-radius:50%;background:#fff;padding:6px;display:block;margin:0 auto;">'
    +   '</div>'
    +   '<div style="font-size:34px;font-weight:900;color:'+heading+';text-align:center;margin:6px 0 18px 0;letter-spacing:.2px;">סיכום משתה פלור - פיטמאסטר אלונים</div>'
    +   '<div style="text-align:center;margin-bottom:14px;"><span style="display:inline-block;font-size:12px;font-weight:700;color:'+accent+';border:1px solid '+accent+';padding:4px 10px;border-radius:999px;letter-spacing:.4px;">דו״ח סיכום משתה</span></div>'
    +   block('תאריך ושעה', p.date + '  ' + p.time)
    +   block('שם המנהל', p.manager)
    +   block('שם הפיט', p.pit)
    +   block('פירוט מנות', coursesText)
    +   block('בעיות במהלך המשתה', p.issuesText)
    +   block('כמות סועדים', p.guestsText)
    +   block('מספרי שולחנות והמלצרים', p.tablesText)
    +   block('דברים שנפתחו', p.openedText)
    +   block('חוסרים', p.missingText)
    +   block('צוות פלור', p.floorText)
    +   block('חוות דעת לקוחות – צוות פלור', p.floorCustomerFeedbackText)
    +   block('צוות מטבח', p.kitchenText)
    +   block('אלכוהול שנמכר', p.alcoholSoldText)
    +   block('אלכוהול שנמזג', p.alcoholPouredText)
    +   block('תקלות בינוי', p.buildFaultsText)
    +   block('התנהלות שוטפי כלים', p.dishwashersConductText)
    +   block('הנחות/OTH', p.discountsText)
    +   block('סיכום הפיטמאסטר', p.pitmasterSummaryText)
    +   block('הערות נוספות', p.extraNotesText)
    + '</div>'
    + '<div style="max-width:800px;margin:0 auto 18px auto;text-align:center;color:#9CA3AF;font-size:12px;">סיכום משתה - פיטמאסטר אלונים</div>'
    + '</div>';

  return { text:text, html:html };
}

function buildPdfHtml_(p, coursesText, logoDataUrl){
  function esc(s){ s = (s==null?'':String(s)); return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  var heading='#111827', bodyColor='#374151', border='#E5E7EB', accent='#D97706';

  function block(title, content){
    return ''
      + '<div style="margin:0 0 16px 0;padding:0 0 10px 0;border-bottom:1px solid '+border+';">'
      +   '<div style="font-size:18px;font-weight:800;color:'+heading+';margin:0 0 6px 0;font-family:Heebo,Arial,Helvetica,sans-serif;text-align:right;">'+ esc(title) +'</div>'
      +   '<div style="font-size:14px;line-height:1.6;color:'+bodyColor+';white-space:pre-wrap;font-family:Heebo,Arial,Helvetica,sans-serif;">'+ esc(content) +'</div>'
      + '</div>';
  }

  var html =
    '<!doctype html><html><head><meta charset="utf-8"><title>סיכום משתה</title>'
  + '<style>html,body{margin:0;padding:0} @page{size:A4;margin:16mm} body{font-family:Heebo,Arial,Helvetica,sans-serif}</style>'
  + '</head><body dir="rtl">'
  +   '<div style="max-width:720px;margin:0 auto;">'
  +     '<div style="text-align:center;margin-bottom:10px;">'
  +       '<img src="'+logoDataUrl+'" style="max-height:90px;border-radius:50%;background:#fff;padding:6px;">'
  +     '</div>'
  +     '<div style="text-align:center;font-size:22px;font-weight:900;color:'+heading+';margin:6px 0 14px 0;">סיכום משתה פלור — פיטמאסטר אלונים</div>'
  +     '<div style="text-align:center;margin-bottom:10px;">'
  +       '<span style="display:inline-block;font-size:11px;font-weight:700;color:'+accent+';border:1px solid '+accent+';padding:3px 8px;border-radius:999px;">דו״ח סיכום משתה</span>'
  +     '</div>'
  +     block('תאריך ושעה', (p.date||'')+'  '+(p.time||''))
  +     block('שם המנהל', p.manager||'')
  +     block('שם הפיט', p.pit||'')
  +     block('פירוט מנות', coursesText||'')
  +     block('בעיות במהלך המשתה', p.issuesText||'')
  +     block('כמות סועדים', p.guestsText||'')
  +     block('מספרי שולחנות והמלצרים', p.tablesText||'')
  +     block('דברים שנפתחו', p.openedText||'')
  +     block('חוסרים', p.missingText||'')
  +     block('צוות פלור', p.floorText||'')
  +     block('חוות דעת לקוחות – צוות פלור', p.floorCustomerFeedbackText||'')
  +     block('צוות מטבח', p.kitchenText||'')
  +     block('אלכוהול שנמכר', p.alcoholSoldText||'')
  +     block('אלכוהול שנמזג', p.alcoholPouredText||'')
  +     block('תקלות בינוי', p.buildFaultsText||'')
  +     block('התנהלות שוטפי כלים', p.dishwashersConductText||'')
  +     block('הנחות/OTH', p.discountsText||'')
  +     block('סיכום הפיטמאסטר', p.pitmasterSummaryText||'')
  +     block('הערות נוספות', p.extraNotesText||'')
  +   '</div>'
  + '</body></html>';

  return html;
}

function createPdfBlob_(p, coursesText){
  var logoBlob = DriveApp.getFileById(FILE_ID).getBlob();
  var logoDataUrl = 'data:'+logoBlob.getContentType()+';base64,'+Utilities.base64Encode(logoBlob.getBytes());
  var pdfHtml = buildPdfHtml_(p, coursesText, logoDataUrl);
  var pdfBlob = Utilities.newBlob(pdfHtml, 'text/html', 'pitmaster-summary.html').getAs('application/pdf');
  var name = 'סיכום משתה אלונים - ' + (p.date || '') + ' ' + (p.time || '') + '.pdf';
  pdfBlob.setName(name);

  var savedFileId = '';
  if (SAVE_PDF_TO_DRIVE && PDF_FOLDER_ID) {
    try {
      var folder = DriveApp.getFolderById(PDF_FOLDER_ID);
      var saved = folder.createFile(pdfBlob);
      savedFileId = saved.getId();
    } catch(e) {
      Logger.log('Save PDF failed: ' + e);
    }
  }
  return { blob: pdfBlob, fileId: savedFileId };
}

function submitForm(p) {
  var lock = LockService.getScriptLock();
  var meta = {
    ua: (p && p.__ua) || '',
    pageUrl: (p && p.__pageUrl) || '',
    date: p && p.date,
    time: p && p.time,
    manager: p && p.manager,
    pit: p && p.pit
  };
  var mailStat = {total: EMAILS.length, ok:0, fail:0, sentList:[], failList:[]};

  try {
    lock.waitLock(30000);

    var cache = CacheService.getScriptCache();
    if (p && p.nonce) {
      var nonceKey = 'nonce_' + p.nonce;
      if (cache.get(nonceKey)) {
        logRun_('חסום כפילות', 'nonce קיים', meta, mailStat);
        return { ok:true, dedup:true };
      }
    }

    var hash = _hashPayload_(p);
    var dupKey = 'dedupe_' + hash;
    if (cache.get(dupKey)) {
      logRun_('חסום כפילות', 'תוכן זהה בתוך חלון זמן', meta, mailStat);
      return { ok:true, dedup:true };
    }

    var v = validatePayload_(p);
    if (!v.ok) {
      logRun_('נחסם', v.reason, meta, mailStat);
      return { ok:false, error: v.reason };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    var HEADERS = [
      'Timestamp',
      'תאריך','שעת משתה','שם המנהל','שם הפיט',
      'פירוט מנות (שורות מרובות)',
      'בעיות במהלך המשתה','כמות סועדים',
      'מספרי שולחנות והמלצרים','דברים שנפתחו',
      'חוסרים','צוות פלור','חוות דעת לקוחות – צוות פלור','צוות מטבח',
      'אלכוהול שנמכר','אלכוהול שנמזג','תקלות בינוי','התנהלות שוטפי כלים',
      'הנחות/OTH','סיכום הפיטמאסטר','הערות נוספות'
    ];
    ensureHeaders_(sh, HEADERS);

    function nv(x){ return (x===undefined || x===null) ? '' : String(x); }

    var courses = Array.isArray(p.courses) ? p.courses : [];
    var coursesText = courses.map(function(c){
      var n = (c && c.name) ? String(c.name) : '';
      var d = (c && c.details) ? String(c.details) : '';
      return n + ' — ' + d;
    }).join('\n');

    sh.appendRow([
      new Date(),
      nv(p.date), nv(p.time), nv(p.manager), nv(p.pit),
      coursesText,
      nv(p.issuesText),
      nv(p.guestsText),
      nv(p.tablesText),
      nv(p.openedText),
      nv(p.missingText),
      nv(p.floorText),
      nv(p.floorCustomerFeedbackText),
      nv(p.kitchenText),
      nv(p.alcoholSoldText),
      nv(p.alcoholPouredText),
      nv(p.buildFaultsText),
      nv(p.dishwashersConductText),
      nv(p.discountsText),
      nv(p.pitmasterSummaryText),
      nv(p.extraNotesText)
    ]);

    var mail = buildMailContent_(p, coursesText);
    var subj = 'סיכום משתה פלור אלונים — ' + nv(p.date) + ' ' + nv(p.time) + ' — ' + nv(p.manager);
    var logoBlob = DriveApp.getFileById(FILE_ID).getBlob();

    var pdfRes = createPdfBlob_(p, coursesText);
    var pdfBlob = pdfRes.blob;
    var pdfSavedId = pdfRes.fileId;

    EMAILS.forEach(function(to){
      if (!to || to.indexOf('@') < 0) return;
      try {
        GmailApp.sendEmail(to.trim(), subj, mail.text, {
          htmlBody: mailhtml,
          inlineImages: { logoImage: logoBlob },
          attachments: [pdfBlob]
        });
        mailStat.ok++; mailStat.sentList.push(to);
      } catch (eSend) {
        mailStat.fail++; mailStat.failList.push(to);
      }
    });

    cache.put(dupKey, '1', 600);
    if (p && p.nonce) cache.put('nonce_'+p.nonce, '1', 600);

    var reasonMain = 'נשלח כולל PDF' + (pdfSavedId ? (' | נשמר בדרייב: ' + pdfSavedId) : '');
    logRun_('הצלחה', reasonMain, meta, mailStat);
    if (mailStat.sentList.length) logMailPerRecipient_(mailStat.sentList, 'נשלח', 'כולל PDF', p);
    if (mailStat.failList.length) logMailPerRecipient_(mailStat.failList, 'נכשל', 'שליחת Gmail נכשלה (עם PDF)', p);

    return { ok:true };
  } catch (err) {
    logRun_('שגיאה', String(err), meta, mailStat);
    if (EMAILS && EMAILS.length) logMailPerRecipient_(EMAILS, 'לא נשלח', 'חריג כללי: '+String(err), p);
    try {
      var subject = 'שגיאה בשליחת סיכום משתה — פיטמאסטר אלונים';
      var bodyTxt =
        'אירעה שגיאה בהגשת טופס.\n\n' +
        'שגיאה: ' + String(err) + '\n' +
        'עמוד: ' + (meta.pageUrl || '') + '\n' +
        'User-Agent: ' + (meta.ua || '') + '\n' +
        'תאריך: ' + (meta.date || '') + ' שעה: ' + (meta.time || '') + '\n' +
        'מנהל: ' + (meta.manager || '') + ' | פיט: ' + (meta.pit || '');
      ERROR_EMAILS.forEach(function(to){ GmailApp.sendEmail(to, subject, bodyTxt); });
    } catch(e2){}
    return { ok:false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function testWrite() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  ensureHeaders_(sh, [
    'Timestamp','תאריך','שעת משתה','שם המנהל','שם הפיט','פירוט מנות (שורות מרובות)',
    'בעיות במהלך המשתה','כמות סועדים','מספרי שולחנות והמלצרים','דברים שנפתחו','חוסרים',
    'צוות פלור','חוות דעת לקוחות – צוות פלור','צוות מטבח','אלכוהול שנמכר','אלכוהול שנמזג',
    'תקלות בינוי','התנהלות שוטפי כלים','הנחות/OTH','סיכום הפיטמאסטר','הערות נוספות'
  ]);
  sh.appendRow(['TEST', new Date(), 'בדיקה חיבור']);
}

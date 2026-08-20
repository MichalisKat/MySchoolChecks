/**
 * MySchool Checks — καταγραφή εγκαταστάσεων/χρήσης.
 *
 * ΟΔΗΓΙΕΣ ΕΓΚΑΤΑΣΤΑΣΗΣ:
 *  1. Πήγαινε στο https://script.google.com → «Νέο έργο».
 *  2. Διάγραψε ό,τι υπάρχει στο Code.gs (μαζί με το προεπιλεγμένο άδειο
 *     "function myFunction() {}" που βάζει αυτόματα η Google) και επικόλλησε
 *     ΟΛΟ αυτό το αρχείο στη θέση του.
 *  3. Το EMAIL_TO παρακάτω είναι ήδη ορισμένο σε katsirntakis@gmail.com.
 *  4. Πάνω δεξιά «Deploy» → «New deployment» → τύπος «Web app».
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  5. Deploy. Θα σου δώσει ένα URL σαν:
 *       https://script.google.com/macros/s/AKfycb.../exec
 *     Αυτό το URL είναι που θα μπει στο main.py (config.py → PING_URL) του
 *     προγράμματος.
 *  6. Κάθε φορά που αλλάζεις τον κώδικα εδώ, πρέπει να κάνεις «Deploy» →
 *     «Manage deployments» → ✏ (Edit) → «New version» → Deploy, ώστε οι
 *     αλλαγές να ισχύσουν στο ΙΔΙΟ URL.
 *  7. Το πρώτο request θα δημιουργήσει αυτόματα ένα Google Sheet
 *     («MySchoolChecks — Εγκαταστάσεις») στο Drive σου, με tab «Εγκαταστάσεις»
 *     (μία γραμμή ανά υπολογιστή) — από εκεί το πρόγραμμα υπολογίζει και το
 *     tab «Χρήστες» που βλέπεις μέσα στην εφαρμογή.
 */

var EMAIL_TO   = 'katsirntakis@gmail.com';
var SHEET_NAME = 'MySchoolChecks — Εγκαταστάσεις';
var SHEET_TAB  = 'Εγκαταστάσεις';

function doGet(e) {
  var action = (e.parameter.action || 'ping').toString().trim();
  try {
    if (action === 'list') {
      return _handleList();
    }
    return _handlePing(e);
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err).setMimeType(ContentService.MimeType.TEXT);
  }
}

function _handlePing(e) {
  var dirName = (e.parameter.dir     || '').toString().trim();
  var dirType = (e.parameter.type    || '').toString().trim(); // "ΠΕ" ή "ΔΕ"
  var version = (e.parameter.version || '').toString().trim();
  var iid     = (e.parameter.iid     || '').toString().trim(); // μοναδικό ανά υπολογιστή

  if (!dirName) {
    return ContentService.createTextOutput('missing dir').setMimeType(ContentService.MimeType.TEXT);
  }

  var sheet = _getSheet();
  var data  = sheet.getDataRange().getValues();
  var now   = new Date();

  // Στήλες: install_id, Διεύθυνση, Τύπος, Αρ. Υπολογιστή, Πρώτη χρήση,
  //         Τελευταία χρήση, Έκδοση
  //
  // ΣΗΜΑΝΤΙΚΟ: ο υπολογιστής αναγνωρίζεται ΑΠΟΚΛΕΙΣΤΙΚΑ από το install_id
  // (μόνιμο, δημιουργείται μία φορά από το ίδιο το πρόγραμμα και επιβιώνει
  // από update/επανεγκατάσταση). ΔΕΝ ψάχνουμε φιλτραρισμένα μέσα στην τρέχουσα
  // Διεύθυνση — αλλιώς, αν επιλεγεί διαφορετική Διεύθυνση σε ένα update,
  // θα δημιουργούνταν διπλή («φανταστική») γραμμή για τον ίδιο υπολογιστή.
  var foundRow = -1;
  if (iid) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === iid) { foundRow = i; break; }
    }
  }

  if (foundRow !== -1) {
    // Ήδη γνωστός υπολογιστής — ενημέρωση στοιχείων + «τελευταία χρήση».
    // (Η Διεύθυνση/Τύπος ενημερώνονται κι αυτά, σε περίπτωση που άλλαξε η
    // επιλογή σε update· ο αύξων αριθμός #Ν παραμένει όπως ήταν.)
    var row = foundRow + 1;
    sheet.getRange(row, 2).setValue(dirName);
    sheet.getRange(row, 3).setValue(dirType);
    sheet.getRange(row, 6).setValue(now);
    if (version) sheet.getRange(row, 7).setValue(version);
  } else {
    // Νέος υπολογιστής (νέο install_id) — υπολογισμός επόμενου αύξοντα
    // αριθμού μέσα στην επιλεγμένη Διεύθυνση.
    var maxSeqForDir = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === dirName && data[i][2] === dirType) {
        var seq = Number(data[i][3]) || 0;
        if (seq > maxSeqForDir) maxSeqForDir = seq;
      }
    }
    var newSeq = maxSeqForDir + 1;
    sheet.appendRow([iid, dirName, dirType, newSeq, now, now, version]);
    var label = dirType + ' ' + dirName + (newSeq > 1 ? ' #' + newSeq : '');
    _sendEmail(label, version, newSeq === 1);
  }

  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}

function _handleList() {
  var sheet = _getSheet();
  var data  = sheet.getDataRange().getValues();
  var counts = {}; // key = type + '|' + dir  →  {dir, type, count}

  for (var i = 1; i < data.length; i++) {
    var dirName = data[i][1], dirType = data[i][2];
    if (!dirName) continue;
    var key = dirType + '|' + dirName;
    if (!counts[key]) counts[key] = {dir: dirName, type: dirType, count: 0};
    counts[key].count++;
  }

  var out = [];
  for (var k in counts) out.push(counts[k]);

  return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
}

function _getSheet() {
  var files = DriveApp.getFilesByName(SHEET_NAME);
  var ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
  }
  var sheet = ss.getSheetByName(SHEET_TAB);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_TAB);
    sheet.appendRow(['install_id', 'Διεύθυνση', 'Τύπος', 'Αρ. Υπολογιστή',
                      'Πρώτη χρήση', 'Τελευταία χρήση', 'Έκδοση']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _sendEmail(label, version, isNewDirectorate) {
  if (!EMAIL_TO) return;
  var subject = isNewDirectorate
    ? 'MySchool Checks — Νέα Διεύθυνση: ' + label
    : 'MySchool Checks — Νέος υπολογιστής: ' + label;
  var body = 'Διεύθυνση: ' + label + '\n' +
             'Έκδοση: ' + (version || '—') + '\n' +
             'Ημερομηνία: ' + new Date().toLocaleString('el-GR');
  MailApp.sendEmail(EMAIL_TO, subject, body);
}

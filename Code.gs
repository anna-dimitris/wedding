/**
 * ════════════════════════════════════════════════════════════════════════
 *  RSVP backend — Save the Date «Άννα & Δημήτρης»
 *  Google Apps Script δεμένο σε ένα Google Sheet.
 *
 *  Τι κάνει:
 *   • Δέχεται POST (JSON) από τη φόρμα του site.
 *   • Γράφει ΜΙΑ γραμμή ανά καλεσμένο (dedup: αν ξαναστείλει, ενημερώνεται).
 *   • Κρατά ζωντανό σύνολο (φύλλο «Σύνολο»): Ναι / Όχι / σύνολο ατόμων.
 *   • Στέλνει email σε κάθε νέα ή ενημερωμένη απάντηση.
 *   • Απορρίπτει bots μέσω honeypot (κρυφό πεδίο «website»).
 *
 *  Ρύθμιση: γράψε το email σου στο NOTIFY_EMAIL, τρέξε μία φορά τη setup(),
 *  και μετά κάνε Deploy → Web app (δες SETUP.md).
 * ════════════════════════════════════════════════════════════════════════
 */

// ⚙️ Email ειδοποιήσεων. Κενό ή "PASTE_EMAIL" = χωρίς αποστολή email.
var NOTIFY_EMAIL = "weddinganndim@gmail.com";

var SHEET_NAME   = "Απαντήσεις";
var SUMMARY_NAME = "Σύνολο";
var HEADERS      = ["Ημ/νία", "Όνομα", "Θα έρθει;", "Άτομα", "Συνοδοί", "Token", "Ενημερώθηκε"];

/** Δέχεται τις υποβολές της φόρμας. */
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);   // απόφυγε ταυτόχρονες εγγραφές
  try{
    var data = {};
    try { data = JSON.parse(e.postData.contents); } catch (err) { data = {}; }

    // honeypot: αν το κρυφό πεδίο έχει τιμή → bot. Αγνόησέ το «σιωπηλά».
    if (data.website){ return ok_({ ok:true, ignored:true }); }

    var name = String(data.name || "").trim();
    if (!name){ return ok_({ ok:false, error:"no name" }); }

    var attending  = String(data.attending  || "").trim();
    var guests     = Number(data.guests) || 0;
    var companions = String(data.companions || "").trim();
    var token      = String(data.token || "").trim();

    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ensureSheet_(ss);
    var now = new Date();

    // dedup: βρες υπάρχουσα εγγραφή (με token, αλλιώς με όνομα)
    var rowIndex = findRow_(sh, token, name);
    // safe_(): εξουδετέρωσε formula-injection (κείμενο που ξεκινά με =,+,-,@) στα πεδία χρήστη
    var row = [now, safe_(name), safe_(attending), guests, safe_(companions), safe_(token), now];
    var isUpdate = false;

    if (rowIndex > 0){
      var firstSeen = sh.getRange(rowIndex, 1).getValue() || now;   // κράτα την αρχική ημ/νία
      row[0] = firstSeen;
      sh.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
      isUpdate = true;
    } else {
      sh.appendRow(row);
    }

    ensureSummary_(ss);
    SpreadsheetApp.flush();               // ώστε το σύνολο να είναι φρέσκο για το email
    notify_(name, attending, guests, companions, isUpdate, ss);
    return ok_({ ok:true, updated:isUpdate });

  } catch (err){
    return ok_({ ok:false, error:String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Άνοιγμα του URL στον browser → μικρό status (για δοκιμή). */
function doGet(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var n  = sh ? Math.max(0, sh.getLastRow() - 1) : 0;
  return ContentService
    .createTextOutput("✅ RSVP endpoint ενεργό. Καταχωρημένες απαντήσεις: " + n)
    .setMimeType(ContentService.MimeType.TEXT);
}

/** Τρέξε ΜΙΑ φορά χειροκίνητα: φτιάχνει τα φύλλα + δίνει άδειες (email). */
function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss);
  ensureSummary_(ss);
  // σβήσε τυχόν κενό προεπιλεγμένο φύλλο («Sheet1» / «Φύλλο1»)
  ss.getSheets().forEach(function(s){
    var nm = s.getName();
    if (nm !== SHEET_NAME && nm !== SUMMARY_NAME && s.getLastRow() === 0){
      try { ss.deleteSheet(s); } catch (e) {}
    }
  });
  SpreadsheetApp.flush();
}

/* ───────────────────────── βοηθητικές ───────────────────────── */

function ensureSheet_(ss){
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0){
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 190); sh.setColumnWidth(5, 220);
  }
  return sh;
}

function ensureSummary_(ss){
  var sm = ss.getSheetByName(SUMMARY_NAME);
  if (!sm){
    sm = ss.insertSheet(SUMMARY_NAME, 0);   // πρώτο φύλλο = ό,τι βλέπει πρώτο το ζευγάρι
    sm.getRange("A1:A5").setValues([
      ["Σύνοψη RSVP"],
      ["Σύνολο απαντήσεων"],
      ["✅ Έρχονται (Ναι)"],
      ["❌ Δεν έρχονται (Όχι)"],
      ["👥 Σύνολο ατόμων που έρχονται"]
    ]);
    sm.getRange("A1:B1").merge().setFontWeight("bold").setFontSize(13);
    sm.getRange("A2:A5").setFontWeight("bold");
    sm.getRange("B2:B5").setFontSize(14).setHorizontalAlignment("right");
    sm.setColumnWidth(1, 240); sm.setColumnWidth(2, 90);
  }
  // Οι φόρμουλες μπαίνουν με setFormulas (ΟΧΙ setValues): το setValues ερμηνεύει το
  // κείμενο με το locale του αρχείου — σε ελληνικό locale το διαχωριστικό είναι «;»
  // και τα COUNTIF/SUMIF με κόμμα γίνονται #ERROR!. Το setFormulas δέχεται πάντα
  // US σύνταξη, σε κάθε locale. Ξαναγράφονται και σε υπάρχον φύλλο, ώστε ένα
  // χαλασμένο «Σύνολο» να αυτοδιορθώνεται στο επόμενο RSVP ή setup().
  var q = "'" + SHEET_NAME + "'";
  sm.getRange("B2:B5").setFormulas([
    ["=COUNTA(" + q + "!B2:B)"],
    ["=COUNTIF(" + q + "!C2:C,\"Ναι\")"],
    ["=COUNTIF(" + q + "!C2:C,\"Όχι\")"],
    ["=SUMIF(" + q + "!C2:C,\"Ναι\"," + q + "!D2:D)"]
  ]);
}

/** Κανονικοποίηση ονόματος για σύγκριση: πεζά, χωρίς τόνους/διαλυτικά, τελικό ς→σ,
 *  ενιαία κενά. Ώστε «ΜΑΡΙΑ» = «Μαρία» και «ΓΙΩΡΓΟΣ» = «Γιώργος» στο dedup. */
function norm_(s){
  return String(s == null ? "" : s).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")   // βγάλε τόνους/διαλυτικά
    .replace(/ς/g, "σ")                        // τελικό σίγμα → σίγμα
    .replace(/\s+/g, " ");                               // ενιαία κενά
}

/** Επιστρέφει τον αριθμό γραμμής υπάρχουσας εγγραφής, ή -1. */
function findRow_(sh, token, name){
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var nName = norm_(name);

  // 1) ταύτιση με token (αν υπάρχει)
  if (token){
    for (var i = 0; i < values.length; i++){
      if (String(values[i][5] || "").trim() === token) return i + 2;
    }
  }
  // 2) fallback: γραμμή ΧΩΡΙΣ token με ίδιο (κανονικοποιημένο) όνομα.
  //    Καλύπτει και το «χωρίς token» flow ΚΑΙ την περίπτωση που ο ίδιος
  //    καλεσμένος έστειλε πρώτα χωρίς token και μετά με token (το token
  //    γράφεται πάνω στην ίδια γραμμή → όχι διπλοεγγραφή).
  for (var j = 0; j < values.length; j++){
    if (!String(values[j][5] || "").trim() && norm_(values[j][1]) === nName) return j + 2;
  }
  return -1;
}

/** Εξουδετερώνει formula/CSV injection: κείμενο που ξεκινά με =,+,-,@,tab,CR
 *  προθεματίζεται με ' ώστε τα Sheets να το κρατούν ως απλό κείμενο (το ' δεν φαίνεται). */
function safe_(s){
  s = String(s == null ? "" : s);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function notify_(name, attending, guests, companions, isUpdate, ss){
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL === "PASTE_EMAIL") return;
  if (MailApp.getRemainingDailyQuota() <= 0) return;   // τελείωσε το ημερήσιο όριο email
  try{
    var sm     = ss.getSheetByName(SUMMARY_NAME);
    var yes    = sm ? sm.getRange("B3").getValue() : "";
    var people = sm ? sm.getRange("B5").getValue() : "";
    var verb   = isUpdate ? "ενημέρωσε" : "έστειλε";
    var subject = "RSVP: " + name + " — " + (attending || ";");
    var body =
      name + " " + verb + " την απάντηση.\n\n" +
      "Θα έρθει: " + (attending || "—") + "\n" +
      "Άτομα: "    + guests + "\n" +
      (companions ? "Συνοδοί: " + companions + "\n" : "") +
      "\n— Σύνολο μέχρι τώρα —\n" +
      "Έρχονται (Ναι): " + yes + "\n" +
      "Άτομα συνολικά: " + people + "\n";
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) { /* μην αποτύχει το request αν χαλάσει το email */ }
}

function ok_(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

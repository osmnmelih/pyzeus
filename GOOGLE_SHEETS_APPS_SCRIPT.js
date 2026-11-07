// =========================================================================
// PYZEUS — automatic data gathering system
// =========================================================================

var SHEET_NAME = "PyZeus_Logs";

var HEADERS = [
  "log_id",
  "student_nickname",
  "timestamp",
  "datetime_local",
  "chapter",
  "topic",
  "activity_type",
  "event_type",
  "concepts_covered",
  "question_number",
  "attempts",
  "first_attempt_correct",
  "hint_used",
  "explanation_viewed",
  "accuracy_pct",
  "tutor_notes"
];

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    // Mevcut saat takip tablosuna dokunmaz, yeni temiz bir sayfa açar
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// -------------------------------------------------------------------------
// 1. POST: Öğrencilerden gelen soru ve sınav loglarını Google Tabloya yazar
// -------------------------------------------------------------------------
function doPost(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      data.log_id || "",
      data.student_nickname || "",
      data.timestamp || new Date().toISOString(),
      data.datetime_local || Utilities.formatDate(new Date(), "Europe/Berlin", "yyyy-MM-dd HH:mm:ss"),
      data.chapter || data.week || "",
      data.topic || "",
      data.activity_type || "Formative Quiz",
      data.event_type || "",
      data.concepts_covered || "",
      data.question_number || "",
      data.attempts || 1,
      data.first_attempt_correct !== undefined ? data.first_attempt_correct : "",
      data.hint_used || false,
      data.explanation_viewed || false,
      data.accuracy_pct || "",
      data.tutor_notes || ""
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------------------------
// 2. GET: Eğitmen Dashboard'una (#admin) tablodaki tüm verileri canlı JSON döner
// -------------------------------------------------------------------------
function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var headers = data[0];
    var logs = [];

    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      logs.push(row);
    }

    return ContentService.createTextOutput(JSON.stringify(logs))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

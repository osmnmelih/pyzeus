// =========================================================================
// PYZEUS — Google Sheets Otomatik Veri Toplayıcı (Apps Script)
// Üniversite: Universität des Saarlandes - Programming for EduTech
// Tutor: Osman Melih Çabık
// =========================================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // Eğer tablo tamamen boşsa (ilk kez çalışıyorsa) başlık satırını ekle
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "log_id",
        "student_nickname",
        "timestamp",
        "datetime_local",
        "week",
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
      ]);
      // Başlık satırını koyu yap ve dondur
      sheet.getRange(1, 1, 1, 16).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    
    // Öğrencinin çözdüğü soru / oturum verisini alt satıra ekle
    sheet.appendRow([
      data.log_id || "",
      data.student_nickname || "",
      data.timestamp || new Date().toISOString(),
      data.datetime_local || Utilities.formatDate(new Date(), "Europe/Berlin", "yyyy-MM-dd HH:mm:ss"),
      data.week || "",
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

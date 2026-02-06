/**
 * BACKEND: SISTEM KAJI SELIDIK IMEN UKM
 * Google Sheet ID: 1mvuTPsthyOFqYsyOKl7a0lQlmRXH-IdSrlDryGpVTbo
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Kajian Kebolehpasaran Program IMEN')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processForm(formObject) {
  try {
    const ss = SpreadsheetApp.openById("1mvuTPsthyOFqYsyOKl7a0lQlmRXH-IdSrlDryGpVTbo");
    let sheet = ss.getSheetByName('Respon_Kajian');
    
    if (!sheet) {
      sheet = ss.insertSheet('Respon_Kajian');
      const headers = [
        "Timestamp", "Email", "Nama", "Umur", "Pekerjaan", "Sektor", 
        "Q1_Sarjana", "Q1_PhD", "Q2_Sarjana", "Q2_PhD", 
        "Q3_Sarjana", "Q3_PhD", "Q4_Sarjana", "Q4_PhD", 
        "Q5_Sarjana", "Q5_PhD", "Q6_Sarjana", "Q6_PhD", 
        "Q7_Sarjana", "Q7_PhD", "Syor_Sarjana", "Syor_PhD", "Cadangan"
      ];
      sheet.appendRow(headers);
    }

    sheet.appendRow([
      new Date(),
      formObject.email,
      formObject.nama || "N/A",
      formObject.umur,
      formObject.pekerjaan === "Lain-lain (Sila nyatakan):" ? "Lain-lain: " + formObject.pekerjaan_lain : formObject.pekerjaan,
      formObject.sektor === "Lain-lain (Sila nyatakan):" ? "Lain-lain: " + formObject.sektor_lain : formObject.sektor,
      formObject.q1_sarjana, formObject.q1_phd,
      formObject.q2_sarjana, formObject.q2_phd,
      formObject.q3_sarjana, formObject.q3_phd,
      formObject.q4_sarjana, formObject.q4_phd,
      formObject.q5_sarjana, formObject.q5_phd,
      formObject.q6_sarjana, formObject.q6_phd,
      formObject.q7_sarjana, formObject.q7_phd,
      formObject.syor_sarjana, formObject.syor_phd,
      formObject.cadangan
    ]);

    return "Success";
  } catch (error) {
    return "Error: " + error.toString();
  }
}

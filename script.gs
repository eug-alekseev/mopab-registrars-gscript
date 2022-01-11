function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Автоматизация');
  menu.addItem('Обработать', 'processAll');
  menu.addToUi();
};

function getAlphabetIndex(c) {
  var A = "A".charCodeAt(0);
  var number = c.charCodeAt(c.length-1) - A;
  if (c.length == 2) {
    number += 26 * (c.charCodeAt(0) - A + 1);
  }
  return number
};

function generateDoc(dir_id, file_name, row, is_finalized) {
  let googleDocTemplate = DriveApp.getFileById(REVIEWED_TEMPLATE_UID);
  if (is_finalized) {
    googleDocTemplate = DriveApp.getFileById(FINALIZED_TEMPLATE_UID);
  }
  
  const destinationFolder = DriveApp.getFolderById(dir_id);
  const doc_copy = googleDocTemplate.makeCopy(file_name, destinationFolder);
  const doc = DocumentApp.openById(doc_copy.getId());
  const body = doc.getBody();

  Object.entries(COLUMNS_CONFIG).forEach(function ([var_name, col]) {
    let value = row[getAlphabetIndex(col)];
    if (var_name === "NSPC_DIAGNOSIS" && value.toLowerCase() === "пересмотр") value = "";
    body.replaceText(`{{${var_name}}}`, value)
  });

  doc.saveAndClose();
  return doc.getUrl();
};

function processAll() {
  SHEETS.forEach(x => processSheet(x));
};

function processSheet(sheet_name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  const rows = sheet.getDataRange().getValues();
  rows.forEach(function(row, index) { 
    if (index === 0) return;  // skip header row
    if (!row[getAlphabetIndex(COLUMNS_CONFIG.READY_TO_SEND)]) return;  // if entry is not ready then skip
    if (row[getAlphabetIndex(COLUMNS_CONFIG.SENT_AT)]) return;  // skip if already sent out
    
    const result_cell = sheet.getRange(index + 1, getAlphabetIndex(COLUMNS_CONFIG.DOC_LINK)+1);
    const status_cell = sheet.getRange(index + 1, getAlphabetIndex(COLUMNS_CONFIG.AUTOMATION_STATUS)+1);
    const sent_at_cell = sheet.getRange(index + 1, getAlphabetIndex(COLUMNS_CONFIG.SENT_AT)+1);
    processRow(row, result_cell, status_cell, sent_at_cell);
  });
};

function processRow(row, result_cell, status_cell, sent_at_cell) {
  const filename = row[getAlphabetIndex(COLUMNS_CONFIG.PATIENT_NAME)];

  const nsps = row[getAlphabetIndex(COLUMNS_CONFIG.NSPC_DIAGNOSIS)];
  let status = row[getAlphabetIndex(COLUMNS_CONFIG.AUTOMATION_STATUS)];

  let dirs = [];
  if (!nsps && status !== "finalized") {
    // review is not needed, save to FINAL
    dirs = DIR_CONFIG.FINAL;
    status_cell.setValue("finalized");
    status = "finalized";
  }
  else if(nsps.toLowerCase() === "пересмотр" && status !== "sent to review") {
    // review is needed, save to REVIEW
    dirs = DIR_CONFIG.REVIEW;
    status_cell.setValue("sent to review");
    status = "sent to review";
  }
  else if(nsps && nsps.toLowerCase() !== "пересмотр" && status !== "review received") {
    // review is ready, regenerate docs to READY
    dirs = DIR_CONFIG.READY;
    status_cell.setValue("review received");
    status = "review received"
  }

  // saving docs
  let doc_url = "";
  dirs.forEach(function(dir) {
    doc_url = generateDoc(dir, filename, row, status === "finalized");
    result_cell.setValue(doc_url);
  });

  // sending docs
  if (["review received", "finalized"].includes(status) && doc_url) {
    const file_id = doc_url.split("id=").slice(-1);
    const gdoc = DriveApp.getFileById(file_id);
    const recipients = EMAILS[row[getAlphabetIndex(COLUMNS_CONFIG.DOCTOR_NAME)]];
    if (recipients) {
      MailApp.sendEmail({
        to: recipients.join(","),
        subject: `Заключение МОПАБ (${row[getAlphabetIndex(COLUMNS_CONFIG.CLINIC_NAME)]}, ${filename})`,
        htmlBody: "",
        attachments: [gdoc.getAs(MimeType.PDF)]
      });
      const current_date = new Date();
      sent_at_cell.setValue(current_date.toLocaleString("ru-RU"));
    }
    else {
      sent_at_cell.setValue(`Адрес не найден!`);
    }
  }
};

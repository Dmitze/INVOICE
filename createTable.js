function createSnapshotSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("А4219");
  const docNumber = sheet.getRange("I11").getValue().toString().trim();
  const rawDate = sheet.getRange("I15").getValue();
  let formattedDate = "";
  if (rawDate instanceof Date) {
    const day = String(rawDate.getDate()).padStart(2, '0');
    const month = String(rawDate.getMonth() + 1).padStart(2, '0');
    const year = rawDate.getFullYear();
    formattedDate = `${day}.${month}.${year}`;
  } else {
    formattedDate = rawDate.toString().trim();
  }

  const subdivision = sheet.getRange("I24:L25").getValues().flat().filter(Boolean).join(" ").trim();
  const rankPib = sheet.getRange("D22:J22").getValues()[0].filter(Boolean).join(" ").trim();
  const totalQty = sheet.getRange("J49").getValue();
  const totalSum = sheet.getRange("K49").getValue();

  const titleParts = [
    docNumber ? `Накладна №${docNumber}` : "",
    formattedDate ? `від ${formattedDate}` : "",
    subdivision ? `(${subdivision})` : "",
    rankPib || "",
    `кількість(${totalQty})`,
    `сума(${totalSum} грн)`
  ];
  const fullTitle = titleParts.filter(Boolean).join(" ").trim() || "Накладна";
  const newSS = SpreadsheetApp.create(fullTitle);
  const newFileId = newSS.getId();
  const critOriginal = ss.getSheetByName("crit");
  const critCopy = critOriginal.copyTo(newSS);
  critCopy.setName("crit");
  critCopy.hideSheet();
  SpreadsheetApp.flush();
  const mainOriginal = ss.getSheetByName("А4219");
  const mainCopy = mainOriginal.copyTo(newSS);
  mainCopy.setName("А4219");
  SpreadsheetApp.flush();
  const defaultSheet = newSS.getSheets()[0];
  newSS.deleteSheet(defaultSheet);
  const lastRow = mainCopy.getLastRow();
  const lastCol = mainCopy.getLastColumn();
  mainCopy.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  return {
    spreadsheet: newSS,
    fileId: newFileId,
    title: fullTitle
  };
}

function exportA4219ToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("А4219");
  const { fileId, title } = createSnapshotSpreadsheet();
  const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const pdfBlob = response.getBlob().setName(`${title}.pdf`);
  const file = DriveApp.createFile(pdfBlob);
  const fileIdFinal = file.getId();
  logExport("PDF", title, fileIdFinal);
  const totalQty = sheet.getRange("J49").getValue();
  const totalSum = sheet.getRange("K49").getValue();
  const summaryHtml = `
    <div style="font-family:Arial; font-size:14px;">
      <p>✅ <b>PDF</b> документ <i>${title}</i> створено.</p>
      <p>📦 Кількість: <b>${totalQty}</b></p>
      <p>💰 Сума: <b>${totalSum}</b></p>
      <p><a href="https://drive.google.com/file/d/${fileIdFinal}" target="_blank">📂 Відкрити документ</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(summaryHtml), "Готово!");
}

function exportA4219ToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("А4219");
  const { fileId, title } = createSnapshotSpreadsheet();
  const excelFile = DriveApp.getFileById(fileId);
  excelFile.setName(`${title}.xlsx`);
  logExport("Excel", title, fileId);
  const totalQty = sheet.getRange("J49").getValue();
  const totalSum = sheet.getRange("K49").getValue();
  const summaryHtml = `
    <div style="font-family:Arial; font-size:14px;">
      <p>✅ <b>Excel</b> документ <i>${title}</i> створено.</p>
      <p>📦 Кількість: <b>${totalQty}</b></p>
      <p>💰 Сума: <b>${totalSum}</b></p>
      <p><a href="https://drive.google.com/file/d/${fileId}" target="_blank">📂 Відкрити документ</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(summaryHtml), "Готово!");
}

function logExport(type, title, fileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Export_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Export_Log");
    logSheet.appendRow(["Тип", "Назва", "Посилання", "Дата"]);
  }
  const url = "https://drive.google.com/file/d/" + fileId;
  logSheet.appendRow([type, title, url, new Date()]);

  const lastRow = logSheet.getLastRow();
  const typeCell = logSheet.getRange(lastRow, 1);
  switch (type) {
    case "PDF":
      typeCell.setBackground("#f89292"); 
      break;
    case "Excel":
      typeCell.setBackground("#06f874"); 
      break;
  }

  logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4)
          .sort({ column: 4, ascending: false });
}

function showLinkModal(type, title, fileId) {
  const url = "https://drive.google.com/file/d/" + fileId;
  const html = `
    <div style="font-family:Arial; font-size:14px; padding:10px;">
      <p>✅ <b>${type}</b> документ <i>${title}</i> успішно створено.</p>
      <p><a href="${url}" target="_blank" style="color:#3367d6;">📂 Відкрити документ</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Експорт завершено");
}

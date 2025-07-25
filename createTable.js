function createSnapshotSpreadsheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("");
  const docNumberRaw = sheet.getRange("I11").getValue().toString().trim();
  const docNumber    = docNumberRaw ? `№${docNumberRaw}` : "";

  const rawDate = sheet.getRange("I15").getValue();
  let formattedDate = "";
  if (rawDate instanceof Date) {
    const d = String(rawDate.getDate()).padStart(2, "0");
    const m = String(rawDate.getMonth() + 1).padStart(2, "0");
    const y = rawDate.getFullYear();
    formattedDate = `від ${d}.${m}.${y}`;
  } else if (rawDate) {
    formattedDate = `від ${rawDate.toString().trim()}`;
  }

  const subdivision = sheet
    .getRange("I24:L25")
    .getValues()
    .flat()
    .filter(Boolean)
    .join(" ")
    .trim();
  const subPart = subdivision ? `(${subdivision})` : "";
  const finder   = sheet.createTextFinder("Всього:").matchCase(false);
  const totalRow = finder.findNext()?.getRow();
  let totalQty = "", totalSum = "";
  if (totalRow) {
    totalQty = sheet.getRange(totalRow, 10).getDisplayValue(); 
    totalSum = sheet.getRange(totalRow, 11).getDisplayValue(); 
  }

  const titleParts = [
    docNumber,
    formattedDate,
    subPart,
    `кількість(${totalQty})`,
    `сума(${totalSum} грн)`
  ];
  const fullTitle = titleParts.filter(Boolean).join(" ").trim() || "Накладна";
  const newSS     = SpreadsheetApp.create(fullTitle);
  const newFileId = newSS.getId();


  ss.getSheetByName("crit")
    .copyTo(newSS)
    .setName("crit")
    .hideSheet();
  SpreadsheetApp.flush();

  ss.getSheetByName("")
    .copyTo(newSS)
    .setName("");
  SpreadsheetApp.flush();
  newSS.deleteSheet(newSS.getSheets()[0]);
  const mainCopy = newSS.getSheetByName("");
  const lr = mainCopy.getLastRow();
  const lc = mainCopy.getLastColumn();
  mainCopy.getRange(1, 1, lr, lc).clearDataValidations();

  return {
    spreadsheet: newSS,
    fileId:      newFileId,
    title:       fullTitle
  };
}


function exportA4219ToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("");
  const { fileId, title, totalQty, totalSum } = createSnapshotSpreadsheet();

  const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}`
    + `/export?format=pdf&portrait=false&fitw=true`
    + `&sheetnames=false&printtitle=false&pagenumbers=false`
    + `&gridlines=false&fzr=false`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const pdfBlob = response.getBlob().setName(`${title}.pdf`);
  const file    = DriveApp.createFile(pdfBlob);
  const fileIdFinal = file.getId();

  logExport("PDF", title, fileIdFinal);

  const summaryHtml = `
    <div style="font-family:Arial; font-size:14px;">
      <p>✅ <b>PDF</b> документ <i>${title}</i> створено.</p>
      <p>📦 Кількість: <b>${totalQty}</b></p>
      <p>💰 Сума: <b>${totalSum}</b></p>
      <p><a href="https://drive.google.com/file/d/${fileIdFinal}" target="_blank">
         📂 Відкрити документ</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(summaryHtml), "Готово!"
  );

  registerDocumentInBook(fileIdFinal, title);
}

function exportA4219ToExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("");
  const { fileId, title, totalQty, totalSum } = createSnapshotSpreadsheet();
  const excelFile = DriveApp.getFileById(fileId);
  excelFile.setName(`${title}.xlsx`);
  logExport("Excel", title, fileId);

  const summaryHtml = `
    <div style="font-family:Arial; font-size:14px;">
      <p>✅ <b>Excel</b> документ <i>${title}</i> створено.</p>
      <p>📦 Кількість: <b>${totalQty}</b></p>
      <p>💰 Сума: <b>${totalSum}</b></p>
      <p><a href="https://drive.google.com/file/d/${fileId}" target="_blank">
         📂 Відкрити документ</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(summaryHtml), "Готово!"
  );

  registerDocumentInBook(fileId, title);
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

function registerDocumentInBook(fileId, title) {
  const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = sourceSS.getSheetByName("");

  const directionValues = sourceSheet.getRange("C20:E20").getValues().flat();
  const direction = directionValues.find(v => v === "Здача" || v === "Видача") || "";
  if (!direction) return;

  const targetSS = SpreadsheetApp.openById("1qUPg_Z2tY5xnou7_RkrADgdoYa0ku4v0I0socFetDKU");
  const targetSheet = targetSS.getSheetByName(direction === "Здача" ? "ЗН" : "ВН");
  const fileUrl = "https://drive.google.com/file/d/" + fileId;
  const existingLinks = targetSheet.getRange("O4:O" + targetSheet.getLastRow()).getValues().flat();
  if (existingLinks.includes(fileUrl)) {
    SpreadsheetApp.getUi().alert(`ℹ️ Документ вже зареєстровано у книзі ${direction}.`);
    return;
  }

  const docDate = sourceSheet.getRange("I15").getValue();
  const day = String(docDate.getDate()).padStart(2, '0');
  const month = String(docDate.getMonth() + 1).padStart(2, '0');
  const year = docDate.getFullYear();
  const formattedDate = `${day}.${month}.${year}`;
  const docNumber = sourceSheet.getRange("I11").getValue().toString().trim();
  const subdivision = sourceSheet.getRange("I24:L25").getValues().flat().filter(Boolean).join(" ").trim();
  const rankPib = sourceSheet.getRange("D22:J22").getValues()[0].filter(Boolean).join(" ").trim();
  const totalQty = sourceSheet.getRange("J49").getValue();
  const g59Value = sourceSheet.getRange("G59").getValue();
  const startRow = 4;
  const dataRange = targetSheet.getRange(startRow, 2, targetSheet.getLastRow() - startRow + 1, 13).getValues();
  let targetRow = startRow;
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i].every(cell => cell === "")) {
      targetRow = startRow + i;
      break;
    }
  }

  const rowValues = [
    formattedDate,            // B
    "Накладна",               // C
    docNumber,                // D
    formattedDate,            // E
    rankPib,                  // F
    subdivision,              // G
    `Кількість: ${totalQty}`,// H
    "", "", "", "", "", ""    // I–N
  ];
  targetSheet.getRange(targetRow, 2, 1, rowValues.length).setValues([rowValues]);
  targetSheet.getRange("F" + targetRow).setValue(2);               // F
  targetSheet.getRange("G" + targetRow).setValue(1);               // G
  targetSheet.getRange("H" + targetRow).setValue(g59Value);        // H
  targetSheet.getRange("O" + targetRow).setValue(fileUrl);         // O

  const allLinks = targetSheet.getRange("O4:O" + targetSheet.getLastRow()).getValues().flat();
  const registeredCount = allLinks.filter(link => typeof link === "string" && link.startsWith("https://")).length;

  SpreadsheetApp.getUi().alert(`✅ Записано в книзі ${direction}\n📌 Рядок №${targetRow}\n🔗 Посилання в O${targetRow}\n📊 Всього зареєстровано: ${registeredCount}`);
}



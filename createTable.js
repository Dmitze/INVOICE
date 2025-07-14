function exportA4219PreservingFormulas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetNamesToCopy = ["crit", "А4219"];
    const availableSheetNames = ss.getSheets().map(sheet => sheet.getName());
    const missingSheets = sheetNamesToCopy.filter(name => !availableSheetNames.includes(name));
    if (missingSheets.length > 0) {
      const message =
        "❌ Не знайдено аркуші:\n" +
        missingSheets.join("\n") +
        "\n\n📋 Доступні аркуші:\n" +
        availableSheetNames.join("\n");
      throw new Error(message);
    }

    const sheet = ss.getSheetByName("А4219");
    const title1 = sheet.getRange("I11").getValue().toString().trim();
    const title2 = sheet.getRange("I15:L15").getValues()[0].join(" ").trim();
    const title3 = sheet.getRange("K49").getValue().toString().trim();
    const title4 = sheet.getRange("D22:J22").getValues()[0].join(" ").trim();
    const fullTitle = [title1, title2, title3, title4].filter(Boolean).join(" – ");
    if (!fullTitle) throw new Error("Назва документа порожня — перевір клітинки.");
    const newSS = SpreadsheetApp.create(fullTitle);
    const newFileId = newSS.getId();
    const critOriginal = ss.getSheetByName("crit");
    const critCopy = critOriginal.copyTo(newSS);
    critCopy.setName("crit");
    critCopy.hideSheet(); // ←
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

    const pdfUrl = `https://docs.google.com/spreadsheets/d/${newFileId}/export?format=pdf&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(pdfUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const pdfBlob = response.getBlob().setName(`${fullTitle}.pdf`);
    DriveApp.createFile(pdfBlob);

    const excelFile = DriveApp.getFileById(newFileId);
    excelFile.setName(`${fullTitle}.xlsx`);
    DriveApp.createFile(excelFile.getBlob());

    SpreadsheetApp.getUi().alert("✅ Готово!\nАркуш");

  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Помилка:\n${error.message}`);
  }
}


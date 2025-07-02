function addProductRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("А4219");
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист "A4219" не знайдено! Перевірте назву листа.');
    return;
  }
  const START_ROW = 29;
  const END_ROW = 49;
  const COLS = 12; // A-L

  let lastRow = START_ROW;
  for (let r = START_ROW; r <= END_ROW; r++) {
    const val = sheet.getRange(r, 2).getValue();
    if (val) lastRow = r;
    else break;
  }

  sheet.insertRowsAfter(lastRow, 1);
  const sourceRange = sheet.getRange(lastRow, 1, 1, COLS);
  const destRange = sheet.getRange(lastRow + 1, 1, 1, COLS);
  sourceRange.copyTo(destRange, {formatOnly: true});
  sourceRange.copyTo(destRange, {validationsOnly: true});
  let row = lastRow + 1;
  sheet.getRange("G" + row).setFormula(`=VLOOKUP(B${row}; 'Довідник'!A22:F; COLUMN(E21); FALSE) & ", " & VLOOKUP(B${row}; 'Довідник'!A22:F; COLUMN(F21); FALSE)`);
  sheet.getRange("H" + row).setFormula(`=IF(B${row}=""; ""; VLOOKUP(B${row}; 'Довідник'!A22:F; COLUMN(C21); FALSE))`);
  sheet.getRange("J" + row).setFormula(`=I${row}`);
  sheet.getRange("K" + row).setFormula(`=J${row}*H${row}`);
}

function deleteProductRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("А4219");
  const START_ROW = 29;
  const COL = 2;
  let lastProductRow = START_ROW;
  for (let r = START_ROW; r <= 49; r++) {
    const val = sheet.getRange(r, COL).getValue();
    if (val) lastProductRow = r;
    else break;
  }
  if (lastProductRow > START_ROW) {
    sheet.deleteRow(lastProductRow);
  } else {
    SpreadsheetApp.getUi().alert("Не можна видалити останню строку!");
  }
}

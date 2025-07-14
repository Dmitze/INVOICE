function copyImportToSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Довідник");
  const targetSheetName = "crit";

  if (!sourceSheet) throw new Error("Аркуш 'Довідник' не знайдено.");

  const range = sourceSheet.getRange("A1").getDataRegion();
  const values = range.getValues();

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) targetSheet = ss.insertSheet(targetSheetName);
  else targetSheet.clear(); 

  targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
}

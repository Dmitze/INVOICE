function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Оновити прописи')
    .addItem('Додати строку', 'addProductRow')
    .addItem('Видалити строку', 'deleteProductRow')
    .addItem('Оновити всі поля', 'updateWordsFieldsDynamic')
    .addToUi();
}

function onEdit(e) {
  const sheetName = "A4219";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  if (!e || !e.range) return;
  const col = e.range.getColumn();
  if (sheet.getName() !== sheetName || !((col === 8) || (col === 11))) return;

  updateWordsFieldsDynamic();
}

function updateWordsFieldsDynamic() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("А4219");
  if (!sheet) return;


  const summaryRow = findSummaryRow(sheet);
  if (!summaryRow) {
    SpreadsheetApp.getUi().alert('Не знайдено рядок "Всього:"');
    return;
  }

  const totalQuantity = sheet.getRange("J" + summaryRow).getValue();
  const totalAmount = sheet.getRange("K" + summaryRow).getValue();
  const transferRow = findRowByText(sheet, "Всього передано");
  const quantityRow = transferRow || (summaryRow + 2);

  if (totalQuantity !== "" && !isNaN(totalQuantity)) {
    const quantityText = numberToWordsUa(totalQuantity);
    sheet.getRange("D" + quantityRow + ":H" + quantityRow).setValues([Array(5).fill(quantityText)]);
  } else {
    sheet.getRange("D" + quantityRow + ":H" + quantityRow).clearContent();
  }

  const amountRow = summaryRow + 3;
  if (totalAmount !== "" && !isNaN(totalAmount)) {
    const amountText = numberToWordsUa(totalAmount);
    sheet.getRange("C" + amountRow + ":H" + amountRow).setValues([Array(6).fill(amountText)]);
    const kopiyky = Math.round((parseFloat(totalAmount) - Math.floor(totalAmount)) * 100);
    let kopiykaWord = kopiykyWordsOnlyUa(kopiyky);
    sheet.getRange("J" + amountRow).setValue(kopiykaWord);
  } else {
    sheet.getRange("C" + amountRow + ":H" + amountRow).clearContent();
    sheet.getRange("J" + amountRow).clearContent();
  }
}

function findRowByText(sheet, needle) {
  const values = sheet.getRange("A1:A1000").getValues();
  for (let i = 0; i < values.length; i++) {
    if ((values[i][0] || "").toString().trim().toLowerCase().indexOf(needle.trim().toLowerCase()) !== -1) {
      return i + 1;
    }
  }
  return null;
}

function findSummaryRow(sheet) {
  return findRowByText(sheet, "Всього:");
}

function numberToWordsUa(number) {
  const units = ['', 'один', 'два', 'три', 'чотири', 'п\'ять', 'шість', 'сім', 'вісім', 'дев\'ять'];
  const unitsF = ['', 'одна', 'дві', 'три', 'чотири', 'п\'ять', 'шість', 'сім', 'вісім', 'дев\'ять'];
  const teens = ['десять', 'одинадцять', 'дванадцять', 'тринадцять', 'чотирнадцять', 'п\'ятнадцять', 'шістнадцять', 'сімнадцять', 'вісімнадцять', 'дев\'ятнадцять'];
  const tens = ['', '', 'двадцять', 'тридцять', 'сорок', 'п\'ятдесят', 'шістдесят', 'сімдесят', 'вісімдесят', 'дев\'яносто'];
  const hundreds = ['', 'сто', 'двісті', 'триста', 'чотириста', 'п\'ятсот', 'шістсот', 'сімсот', 'вісімсот', 'дев\'ятсот'];

  function getPlural(number, forms) {
    if (!forms || forms.length !== 3) throw new Error('forms argument must be an array of three strings');
    number = Math.abs(number) % 100;
    const n = number % 10;
    if (number >= 11 && number <= 19) return forms[2];
    if (n === 1) return forms[0];
    if (n >= 2 && n <= 4) return forms[1];
    return forms[2];
  }

  function convertGroup(num, isThousand) {
    let result = '';
    const h = Math.floor(num / 100);
    if (h > 0) result += hundreds[h] + ' ';
    const t = Math.floor((num % 100) / 10);
    const u = num % 10;
    if (t === 1 && u !== 0) {
      result += teens[u] + ' ';
    } else {
      if (t > 0) result += tens[t] + ' ';
      if (u > 0) result += (isThousand ? unitsF[u] : units[u]) + ' ';
    }
    return result.trim();
  }
  number = parseFloat(number).toFixed(2);
  const integerPart = Math.floor(parseFloat(number));
  if (integerPart === 0) return 'нуль';
  let result = '';
  const million = Math.floor(integerPart / 1000000);
  const thousand = Math.floor((integerPart / 1000) % 1000);
  const unit = integerPart % 1000;

  if (million > 0) {
    result += convertGroup(million, false) + ' ' + getPlural(million, ['мільйон', 'мільйона', 'мільйонів']) + ' ';
  }
  if (thousand > 0) {
    result += convertGroup(thousand, true) + ' ' + getPlural(thousand, ['тисяча', 'тисячі', 'тисяч']) + ' ';
  }
  if (unit > 0) {
    result += convertGroup(unit, false) + ' ';
  }
  return result.trim();
}

function kopiykyWordsOnlyUa(number) {
  const unitsF = ['нуль', 'одна', 'дві', 'три', 'чотири', 'п\'ять', 'шість', 'сім', 'вісім', 'дев\'ять'];
  const teens = ['десять', 'одинадцять', 'дванадцять', 'тринадцять', 'чотирнадцять', 'п\'ятнадцять', 'шістнадцять', 'сімнадцять', 'вісімнадцять', 'дев\'ятнадцять'];
  const tens = ['', '', 'двадцять', 'тридцять', 'сорок', 'п\'ятдесят', 'шістдесят', 'сімдесят', 'вісімдесят', 'дев\'яносто'];

  number = Number(number);
  let word = '';
  if (number === 0) {
    word = 'нуль';
  } else if (number > 9 && number < 20) {
    word = teens[number - 10];
  } else {
    let t = Math.floor(number / 10);
    let u = number % 10;
    if (t > 0) word += tens[t] + ' ';
    word += unitsF[u];
  }
  return word.trim();
}


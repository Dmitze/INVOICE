function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️Меню')
    .addItem('➤Додати строку', 'addProductRow')
    .addItem('➤Видалити строку', 'deleteProductRow')
    .addItem('➤Розрахувати загальну суму', 'updateWordsFieldsDynamic')
    .addToUi();
}

function onEdit(e) {
  const sheetName = "";
  const dictSheetName = "";
  const mvoSheetName = "";
  const categoryColumn = 7; // G
  const itemColumn = 2;     // B
  const valueColumn = 9;    // I
  const firstRow = 29;
  const lastRow = 48;
  const contactEmail = "nrs.a4219@gmail.com";
  const PIB_AND_RANK_CELL = "G59";

  if (e && e.range && e.range.getSheet().getName() === sheetName) {
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    const sheet = e.range.getSheet();
    if (
      row >= firstRow &&
      row <= lastRow &&
      col === valueColumn
    ) {
      const itemName = sheet.getRange(row, itemColumn).getValue();
      const category = sheet.getRange(row, categoryColumn).getValue();
      const inputValue = range.getValue();
      if (!itemName || !category || inputValue === "") return;
      const dictSheet = e.source.getSheetByName(dictSheetName);
      const dictData = dictSheet.getRange(2, 1, dictSheet.getLastRow() - 1, 6).getValues();
      let maxAllowed = null;
      let categoryLabel = "";
      let dictColumn = "";
      for (let i = 0; i < dictData.length; i++) {
        if (dictData[i][0] === itemName) {
          if (category === "І") {
            maxAllowed = dictData[i][4];
            categoryLabel = "Категорія 1";
            dictColumn = "E";
          } else if (category === "ІІ") {
            maxAllowed = dictData[i][5];
            categoryLabel = "Категорія 2";
            dictColumn = "F";
          }
          break;
        }
      }

      function showError(message) {
        SpreadsheetApp.getUi().alert(
          "Шановний\n\n" +
            message +
            "\n\nЩо робити: Перевірте правильність вибору категорії й найменування, а також зверніться до відповідального за ведення у таблиці Речовий склад.\n" +
            `Контакт для звернень: ${contactEmail}\n` +
            `Деталі: шукалось значення у таблиці Речовий склад!${dictColumn} на майно "${itemName}" та категорії "${categoryLabel}".`
        );
      }
      if (maxAllowed === null || maxAllowed === "" || Number(maxAllowed) === 0) {
        showError(
          `${categoryLabel}: значення відсутнє у таблиці для "${itemName}".\n` +
            "Введення кількості неможливе. Поле буде очищено."
        );
        range.setValue("");
        return;
      }
      if (Number(inputValue) > Number(maxAllowed)) {
        showError(
          `Ви не можете ввести більше ніж ${maxAllowed} для "${itemName}" (${categoryLabel}).\n` +
            `Максимально дозволено згідно довідника — ${maxAllowed}. Значення буде автоматично виправлене.`
        );
        range.setValue(maxAllowed);
        return;
      }
    }
    if (col === 8 || col === 11) {
      if (typeof updateWordsFieldsDynamic === "function") {
        updateWordsFieldsDynamic();
      }
    }
    const menuRange = sheet.getRange("I24:L25");
    const longHeight = 131; 
    const defaultHeight = 21;  
    const longTextLength = 50; 

    if (
      menuRange.getRow() <= row && row <= menuRange.getLastRow() &&
      menuRange.getColumn() <= col && col <= menuRange.getLastColumn()
    ) {
      const value = range.getValue();
      if (typeof value === 'string' && value.length > longTextLength) {
        sheet.setRowHeight(row, longHeight);
      } else {
        sheet.setRowHeight(row, defaultHeight);
      }

      const selectedSubdivision = value;
      if (!selectedSubdivision) {
        sheet.getRange(PIB_AND_RANK_CELL).setValue("");
        return;
      }

      const mvoSheet = e.source.getSheetByName(mvoSheetName);
      if (!mvoSheet) {
        sheet.getRange(PIB_AND_RANK_CELL).setValue("");
        return;
      }

      const lastRowMVO = mvoSheet.getLastRow();
      const subList = mvoSheet.getRange(2, 4, lastRowMVO - 1, 1).getValues().flat();
      const rankList = mvoSheet.getRange(2, 2, lastRowMVO - 1, 1).getValues().flat();
      const pibList = mvoSheet.getRange(2, 3, lastRowMVO - 1, 1).getValues().flat();
      const idx = subList.findIndex(v => v === selectedSubdivision);

      if (idx !== -1) {
        const rank = rankList[idx] || "";
        const pib = pibList[idx] || "";
        const pibParts = pib.trim().split(" ");
        let shortPib = pib;
        if (pibParts.length >= 2) {
          const lastName = pibParts[0];
          const firstName = pibParts[1];
          const firstInitial = firstName ? (firstName[0] + ".") : "";
          shortPib = `${firstInitial} ${lastName}`;
        }
        const result = `${rank} ${shortPib}`.trim();
        sheet.getRange(PIB_AND_RANK_CELL).setValue(result);
      } else {
        sheet.getRange(PIB_AND_RANK_CELL).setValue("");
      }
    }
  }
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️Меню')
    .addItem('➤Додати строку', 'addProductRow')
    .addItem('➤Видалити строку', 'deleteProductRow')
    .addItem('➤Розрахувати загальну суму', 'updateWordsFieldsDynamic')
   .addItem('➤ Створити копію шаблону', 'exportA4219PreservingFormulas')
    .addToUi();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ Меню")
    .addSubMenu(
      ui.createMenu("📦 Управління позиціями")
        .addItem("➕ Додати строку", "addProductRow")
        .addItem("➖ Видалити строку", "deleteProductRow")
    )
    .addItem("💰 Розрахувати загальну суму", "updateWordsFieldsDynamic")
    .addSubMenu(
      ui.createMenu("📤 Експорт")
        .addItem("📄 Створити PDF", "exportA4219ToPDF")
        .addItem("📊 Створити Excel", "exportA4219ToExcel")
    )
    .addToUi();
}

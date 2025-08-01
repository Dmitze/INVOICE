# INVOICE Google Apps Script

**INVOICE** — це набір Google Apps Script для повної автоматизації роботи з інвентаризаційними накладними в Google Таблицях. Скрипти:

- Ведуть облік майна та контролюють ліміти по категоріях  
- Автоматично конвертують кількість і суму в український текст  
- Дозволяють експортувати накладні в PDF та Excel  
- Реєструють створені файли в “Книзі реєстрації та руху облікових документів”  
- Підставляють звання, ПІБ та підрозділи відповідальних осіб  
- Додають зручне кастомне меню для управління  

Для коректної роботи ви маєте мати три додаткові аркуші:

- **Речовий склад** — довідник майна  
- **ШПС** — служба планово-складських залишків  
- **Книга реєстрації та руху облікових документів** — журнал видачі/здачі 

---

## Основні можливості

- Додавання та видалення рядків товарів  
- Перевірка введеної кількості за довідником  
- Конвертація чисел і копійок у слова українською  
- Генерація PDF та Excel з автоматичною реєстрацією  
- Автоматичне оновлення ПІБ, звань та підрозділів  
- Збереження історії експортів у “Export_Log”  
- Зручне меню “⚙️ Меню” в інтерфейсі  

---

## Огляд скриптів

- **add.js**  
  - `addProductRow()` — додає новий рядок товару, копіює форматування та формули.
  - `deleteProductRow()` — видаляє останній заповнений рядок товару.

- **createTable.js**  
  - `createSnapshotSpreadsheet()` — створює копію накладної та довідника для експорту.
  - `exportToPDF()` — експортує накладну у PDF, записує інформацію у лог.
  - `exportToExcel()` — експортує накладну у Excel.
  - `logExport()` — веде журнал експортованих документів.
  - `registerDocumentInBook()` — реєструє створений документ у книзі видачі/здачі.

- **ImportToSnapshot.js**  
  - `copyImportToSnapshot()` — копіює дані з листа "Довідник" у технічний лист "crit", який використовується для експорту.

- **numberToWordsUa.js**  
  - `onEdit(e)` — основна функція-обробник змін:  
    - Контролює ліміти по кількості.
    - Оновлює суму та кількість у словах.
    - Автоматично підбирає відповідальних осіб по підрозділах.
  - `updateWordsFieldsDynamic()` — оновлює текстовий запис суми та кількості.
  - `numberToWordsUa(number)` — переводить число у слова (українською).
  - `kopiykyWordsOnlyUa(number)` — переводить копійки у слова.
  - Допоміжні функції для пошуку та заповнення даних.

- **menu.js**  
  - `onOpen()` — додає кастомне меню до Google Таблиці для зручного доступу до всіх функцій.

---

## Як почати користуватись

1. **Створіть Google Таблицю** із такими листами:
    - `Накладна` — основна накладна
    - `Довідник` — довідник майна
    - `МВО` — список відповідальних осіб та підрозділів
    - `crit` — технічний лист (створюється автоматично)
    - `Export_Log` — журнал експортів (створюється автоматично)

2. **Додайте скрипти** у редактор Apps Script (кожен файл окремо).

3. **Надайте необхідні дозволи** (Google Drive, таблиці).

4. **Відкрийте таблицю** — з’явиться меню "⚙️ Меню".

5. **Користуйтеся меню** для додавання/видалення рядків, розрахунків та експорту.

---

## Приклад структури проекту

```
INVOICE/
├── add.js
├── createTable.js
├── ImportToSnapshot.js
├── menu.js
├── numberToWordsUa.js
├── README.md
├── LICENSE
└── CONTRIBUTING.md
```

---

## Контакт

З питань по роботі або пропозицій:  
- Мій профіль: https://github.com/Dmitze  
---

## Проєкт та📄 Ліцензія

- Репозиторій: https://github.com/Dmitze/INVOICE  
- Ліцензія MIT: https://github.com/Dmitze/INVOICE/blob/main/LICENSE  

---




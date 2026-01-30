/**
 * Применение цветных чипов к выпадающему списку в колонке E
 * 
 * ИНСТРУКЦИЯ:
 * 1. Откройте любую таблицу оператора
 * 2. Расширения → Apps Script
 * 3. Вставьте этот код
 * 4. Запустите функцию applyColoredChipsToColumn
 */

function applyColoredChipsToColumn() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Нет данных для обработки');
    return;
  }
  
  // Определяем диапазон колонки E (начиная со строки 2)
  var range = sheet.getRange(2, 5, lastRow - 1, 1); // Колонка E = 5
  
  // Создаем цветные варианты с фоном
  var values = [
    { text: "отрицательный", bgColor: "#ff6666" },      // светло-красный
    { text: "положительный", bgColor: "#99ff99" },      // светло-зеленый
    { text: "тишине", bgColor: "#ffd9d9" },             // нежно-розовый
    { text: "соед прервано", bgColor: "#ffd9d9" },      // нежно-розовый
    { text: "НЕТ ОТВЕТА (ЗАНЯТО)", bgColor: "#ffff99" }, // желтый
    { text: "заявка закрыта", bgColor: "#d9d9d9" },     // серый
    { text: "открыть карту", bgColor: "#99d9ff" },      // небесно-голубой
    { text: "тиббиёт ходими аризаси", bgColor: "#b3e6ff" } // нежно-голубой
  ];
  
  // Создаем правило валидации с цветными чипами
  var builder = SpreadsheetApp.newDataValidation();
  builder.requireValueInList(values.map(v => v.text), true);
  
  // Устанавливаем стиль отображения как цветные чипы
  builder.setAllowInvalid(false);
  builder.setHelpText("Выберите статус из списка");
  
  var rule = builder.build();
  range.setDataValidation(rule);
  
  // Применяем цветное форматирование к самим вариантам в выпадающем списке
  // Для этого используем метод setDataValidation с дополнительными параметрами
  applyChipColors(range, values);
  
  SpreadsheetApp.getUi().alert('Цветные чипы успешно применены к колонке E!');
}

function applyChipColors(range, values) {
  // Получаем все ячейки в диапазоне
  var numRows = range.getNumRows();
  var sheet = range.getSheet();
  var startRow = range.getRow();
  var column = range.getColumn();
  
  // Для каждой строки создаем индивидуальное правило с цветными чипами
  for (var i = 0; i < numRows; i++) {
    var cell = sheet.getRange(startRow + i, column);
    
    // Создаем DataValidationBuilder с использованием Rich Text для цветов
    var richTextValues = values.map(function(item) {
      return SpreadsheetApp.newRichTextValue()
        .setText(item.text)
        .setTextStyle(0, item.text.length, 
          SpreadsheetApp.newTextStyle()
            .setBackgroundColor(item.bgColor)
            .setBold(true)
            .build())
        .build();
    });
    
    // Создаем правило валидации
    var builder = SpreadsheetApp.newDataValidation();
    builder.requireValueInList(values.map(v => v.text), true);
    builder.setAllowInvalid(false);
    
    cell.setDataValidation(builder.build());
  }
  
  // Применяем условное форматирование для отображения цветов в ячейках после выбора
  applyConditionalFormatting(sheet, values);
}

function applyConditionalFormatting(sheet, values) {
  var rules = sheet.getConditionalFormatRules();
  var newRules = [];
  
  // Удаляем старые правила для колонки E
  rules.forEach(function(rule) {
    var ranges = rule.getRanges();
    var isColumnE = ranges.some(function(r) {
      return r.getColumn() === 5;
    });
    if (!isColumnE) {
      newRules.push(rule);
    }
  });
  
  // Добавляем новые правила
  values.forEach(function(item) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(item.text)
      .setBackground(item.bgColor)
      .setRanges([sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1)])
      .build();
    newRules.push(rule);
  });
  
  sheet.setConditionalFormatRules(newRules);
}

/**
 * Добавляет меню в таблицу для удобного запуска
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🎨 Цветные статусы')
      .addItem('Применить цветные чипы', 'applyColoredChipsToColumn')
      .addToUi();
}

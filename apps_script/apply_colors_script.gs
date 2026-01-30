/**
 * Применяет цветовое форматирование к колонке E (Статусы)
 * Версия: 3.2 (светлая цветовая гамма)
 * 
 * ЦВЕТА:
 * • Отрицательный → Светло-красный
 * • Положительный → Светло-зеленый
 * • Тишине → Нежно-розовый
 * • Соед прервано → Нежно-розовый
 * • НЕТ ОТВЕТА (ЗАНЯТО) → Светло-желтый
 * • Заявка закрыта → Светло-серый
 * • Открыть карту → Светло-небесный
 * • Тиббиёт ходими аризаси → Нежно-голубой
 */
function applyStatusColorFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FIKSA");
  
  if (!sheet) {
    Logger.log("Лист FIKSA не найден");
    SpreadsheetApp.getUi().alert("Лист FIKSA не найден!");
    return;
  }
  
  var lastRow = sheet.getMaxRows();
  if (lastRow < 1000) lastRow = 1000;
  
  // Удаляем старые правила условного форматирования для колонки E
  var rules = sheet.getConditionalFormatRules();
  var newRules = [];
  
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var keepRule = true;
    
    for (var j = 0; j < ranges.length; j++) {
      if (ranges[j].getColumn() === 5) {
        keepRule = false;
        break;
      }
    }
    
    if (keepRule) {
      newRules.push(rules[i]);
    }
  }
  
  var range = sheet.getRange("E2:E" + lastRow);
  
  // Создаем правила для каждого статуса

  // отрицательный
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("отрицательный")
    .setBackground("#ff6666")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // положительный
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("положительный")
    .setBackground("#99ff99")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // тишине
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("тишине")
    .setBackground("#ffd9d9")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // соед прервано
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("соед прервано")
    .setBackground("#ffd9d9")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // НЕТ ОТВЕТА (ЗАНЯТО)
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("НЕТ ОТВЕТА (ЗАНЯТО)")
    .setBackground("#ffff99")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // заявка закрыта (не удалось дозвониться)
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("заявка закрыта (не удалось дозвониться)")
    .setBackground("#d9d9d9")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // открыть карту
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("открыть карту")
    .setBackground("#99d9ff")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // тиббиёт ходими аризаси
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("тиббиёт ходими аризаси")
    .setBackground("#b3e6ff")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule);

  // Применяем все правила
  sheet.setConditionalFormatRules(newRules);
  
  // Создаем выпадающий список со всеми статусами
  var statusList = [
    "отрицательный",
    "положительный",
    "тишине",
    "соед прервано",
    "НЕТ ОТВЕТА (ЗАНЯТО)",
    "заявка закрыта (не удалось дозвониться)",
    "открыть карту",
    "тиббиёт ходими аризаси"
  ];
  
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusList, true)
    .setAllowInvalid(false)
    .build();
  
  range.setDataValidation(rule);
  
  Logger.log("✓ Цветовое форматирование и выпадающий список применены к колонке E");
  SpreadsheetApp.getUi().alert(
    "✅ Форматирование применено!\\n\\n" +
    "Колонка E теперь с:\\n" +
    "• Выпадающим списком статусов\\n" +
    "• Цветной подсветкой для каждого статуса\\n\\n" +
    "Всего правил: " + newRules.length
  );
}

/**
 * Создает меню для быстрого доступа
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🎨 Цвета")
    .addItem("Применить цвета к статусам", "applyStatusColorFormatting")
    .addToUi();
}

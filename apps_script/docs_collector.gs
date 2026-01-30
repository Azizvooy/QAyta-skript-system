/**
 * =============================================================================
 * СБОРЩИК ДАННЫХ В GOOGLE DOCS
 * =============================================================================
 * Версия: 1.0
 * Дата: 01.12.2025
 * 
 * 📋 НАЗНАЧЕНИЕ:
 * Собирает все данные из таблиц операторов и записывает в единый Google Docs
 * документ в структурированном формате для последующей обработки через Python
 * 
 * 🔄 ПРОЦЕСС:
 * 1. Apps Script собирает данные из всех таблиц → записывает в Google Docs
 * 2. Python читает Google Docs через API → обрабатывает данные
 * 3. Python записывает результаты обратно в Google Sheets
 * 
 * 📝 ФОРМАТ ДАННЫХ В DOCS:
 * JSON Lines (по одной записи на строку):
 * {"operator":"ФИО","date":"01.12.2024","card":"1234","status":"положительный"}
 * 
 * 🚀 ИСПОЛЬЗОВАНИЕ:
 * 1. Создайте новый Google Docs документ
 * 2. Скопируйте ID документа
 * 3. Вставьте ID в константу DOCS_ID ниже
 * 4. Запустите: Меню → Собрать все данные в Docs
 * =============================================================================
 */

// =============================================================================
// НАСТРОЙКИ
// =============================================================================

// ID документа Google Docs куда будут собираться данные
var DOCS_ID = "ВСТАВЬТЕ_ID_ДОКУМЕНТА_СЮДА";

// ID таблицы со списком операторов (текущая таблица)
var SETTINGS_SHEET_NAME = "Настройки";

// Максимальное количество записей за один запуск (чтобы не превысить лимиты)
var MAX_RECORDS_PER_RUN = 10000;

// =============================================================================
// МЕНЮ
// =============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("📄 Docs Collector")
    .addItem("🔄 Собрать все данные в Docs", "collectAllDataToDocs")
    .addItem("🗑️ Очистить документ Docs", "clearDocsDocument")
    .addSeparator()
    .addItem("📊 Собрать архивные данные", "collectArchiveDataToDocs")
    .addItem("📈 Собрать текущую статистику", "collectCurrentStatsToDocs")
    .addSeparator()
    .addItem("ℹ️ Показать инструкцию", "showInstructions")
    .addToUi();
}

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Собирает все данные из всех таблиц операторов в Google Docs
 */
function collectAllDataToDocs() {
  var startTime = new Date().getTime();
  
  Logger.log("========================================");
  Logger.log("НАЧАЛО СБОРА ДАННЫХ В GOOGLE DOCS");
  Logger.log("========================================\n");
  
  // Проверка настроек
  if (DOCS_ID === "ВСТАВЬТЕ_ID_ДОКУМЕНТА_СЮДА") {
    SpreadsheetApp.getUi().alert(
      "⚠️ Ошибка конфигурации\n\n" +
      "Необходимо:\n" +
      "1. Создать новый Google Docs документ\n" +
      "2. Скопировать ID из URL документа\n" +
      "3. Вставить ID в константу DOCS_ID в коде скрипта"
    );
    return;
  }
  
  try {
    // Получаем список операторов
    var operators = getOperatorList();
    Logger.log("Найдено операторов: " + operators.length);
    
    if (operators.length === 0) {
      throw new Error("Нет операторов в листе 'Настройки'");
    }
    
    // Открываем документ
    var doc = DocumentApp.openById(DOCS_ID);
    var body = doc.getBody();
    
    // Очищаем документ
    body.clear();
    
    // Добавляем заголовок и метаданные
    var header = body.appendParagraph("АРХИВ ДАННЫХ ОПЕРАТОРОВ");
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    body.appendParagraph("Дата сбора: " + new Date().toLocaleString());
    body.appendParagraph("Количество операторов: " + operators.length);
    body.appendParagraph("Формат: JSON Lines (по одной записи на строку)");
    body.appendParagraph("=" .repeat(80));
    body.appendParagraph("");
    
    // Собираем данные
    var totalRecords = 0;
    var processedOperators = 0;
    
    for (var i = 0; i < operators.length; i++) {
      var op = operators[i];
      
      if (op.status.toLowerCase() !== "активен") {
        Logger.log("Пропускаем неактивного оператора: " + op.name);
        continue;
      }
      
      Logger.log("\n▶ Обработка: " + op.name + " (" + (i+1) + "/" + operators.length + ")");
      
      try {
        var records = collectOperatorData(op);
        
        if (records > 0) {
          totalRecords += records;
          processedOperators++;
          Logger.log("  ✓ Собрано записей: " + records);
        } else {
          Logger.log("  ⊗ Нет данных");
        }
        
        // Проверка лимита
        if (totalRecords >= MAX_RECORDS_PER_RUN) {
          Logger.log("\n⚠️ Достигнут лимит записей: " + MAX_RECORDS_PER_RUN);
          Logger.log("Остановка сбора. Обработано операторов: " + processedOperators);
          break;
        }
        
      } catch (err) {
        Logger.log("  ✗ Ошибка для " + op.name + ": " + err.message);
      }
    }
    
    // Добавляем футер
    body.appendParagraph("");
    body.appendParagraph("=" .repeat(80));
    body.appendParagraph("ИТОГО:");
    body.appendParagraph("Обработано операторов: " + processedOperators);
    body.appendParagraph("Всего записей: " + totalRecords);
    body.appendParagraph("Время сбора: " + Math.round((new Date().getTime() - startTime) / 1000) + " сек");
    
    var duration = Math.round((new Date().getTime() - startTime) / 1000);
    
    Logger.log("\n========================================");
    Logger.log("✅ СБОР ЗАВЕРШЕН");
    Logger.log("Операторов: " + processedOperators);
    Logger.log("Записей: " + totalRecords);
    Logger.log("Время: " + duration + " сек");
    Logger.log("========================================");
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "✅ Данные собраны!\n\n" +
      "Операторов: " + processedOperators + "\n" +
      "Записей: " + totalRecords + "\n" +
      "Время: " + duration + " сек\n\n" +
      "Документ: " + doc.getName(),
      "Готово",
      10
    );
    
  } catch (err) {
    Logger.log("\n❌ КРИТИЧЕСКАЯ ОШИБКА: " + err);
    Logger.log("Stack: " + err.stack);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "❌ Ошибка: " + err.message,
      "Ошибка",
      10
    );
  }
}

/**
 * Собирает данные одного оператора из всех архивных листов
 */
function collectOperatorData(operator) {
  var doc = DocumentApp.openById(DOCS_ID);
  var body = doc.getBody();
  
  // Открываем таблицу оператора
  var spreadsheet = SpreadsheetApp.openById(operator.spreadsheetId);
  var sheets = spreadsheet.getSheets();
  
  var totalRecords = 0;
  
  // Проходим по всем листам
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Пропускаем служебные листы
    if (sheetName === "Статистика" || 
        sheetName === "Предыдущий месяц" || 
        sheetName === "Сводка по дням" || 
        sheetName === "Настройки") {
      continue;
    }
    
    // Читаем данные с листа
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;
    
    Logger.log("    Лист: " + sheetName + " (" + (lastRow - 1) + " строк)");
    
    // Читаем колонки B-I (номер карты, данные, статус, дата)
    var data = sheet.getRange(2, 2, lastRow - 1, 8).getValues();
    
    var recordsFromSheet = 0;
    
    for (var j = 0; j < data.length; j++) {
      var cardNum = String(data[j][0] || "").trim();    // B - номер карты
      var status = String(data[j][3] || "").trim();     // E - статус
      var dateValue = data[j][7];                       // I - дата
      
      if (!cardNum) continue;
      
      // Парсим дату
      var dateStr = formatDate(dateValue);
      if (!dateStr) continue;
      
      // Формируем JSON запись
      var record = {
        operator: operator.name,
        sheet: sheetName,
        card: cardNum,
        status: status,
        date: dateStr
      };
      
      // Записываем в документ
      body.appendParagraph(JSON.stringify(record));
      
      recordsFromSheet++;
      totalRecords++;
      
      // Проверка лимита
      if (totalRecords >= MAX_RECORDS_PER_RUN) {
        break;
      }
    }
    
    Logger.log("      Записей: " + recordsFromSheet);
    
    if (totalRecords >= MAX_RECORDS_PER_RUN) {
      break;
    }
  }
  
  return totalRecords;
}

/**
 * Собирает только архивные данные (все листы кроме Статистика и Предыдущий месяц)
 */
function collectArchiveDataToDocs() {
  Logger.log("Сбор АРХИВНЫХ данных...");
  collectAllDataToDocs();
}

/**
 * Собирает только текущую статистику (листы Статистика и Предыдущий месяц)
 */
function collectCurrentStatsToDocs() {
  var startTime = new Date().getTime();
  
  Logger.log("========================================");
  Logger.log("СБОР ТЕКУЩЕЙ СТАТИСТИКИ В DOCS");
  Logger.log("========================================\n");
  
  if (DOCS_ID === "ВСТАВЬТЕ_ID_ДОКУМЕНТА_СЮДА") {
    SpreadsheetApp.getUi().alert("⚠️ Необходимо настроить DOCS_ID");
    return;
  }
  
  try {
    var operators = getOperatorList();
    var doc = DocumentApp.openById(DOCS_ID);
    var body = doc.getBody();
    
    body.clear();
    
    var header = body.appendParagraph("ТЕКУЩАЯ СТАТИСТИКА ОПЕРАТОРОВ");
    header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    body.appendParagraph("Дата сбора: " + new Date().toLocaleString());
    body.appendParagraph("=" .repeat(80));
    body.appendParagraph("");
    
    var totalRecords = 0;
    
    for (var i = 0; i < operators.length; i++) {
      var op = operators[i];
      if (op.status.toLowerCase() !== "активен") continue;
      
      Logger.log("▶ " + op.name);
      
      try {
        var spreadsheet = SpreadsheetApp.openById(op.spreadsheetId);
        
        // Собираем из листа "Статистика"
        var statsSheet = spreadsheet.getSheetByName("Статистика");
        if (statsSheet) {
          var stats = collectStatsFromSheet(statsSheet, op.name, "Текущий месяц");
          if (stats) {
            body.appendParagraph(JSON.stringify(stats));
            totalRecords++;
          }
        }
        
        // Собираем из листа "Предыдущий месяц"
        var prevSheet = spreadsheet.getSheetByName("Предыдущий месяц");
        if (prevSheet) {
          var prevStats = collectStatsFromSheet(prevSheet, op.name, "Предыдущий месяц");
          if (prevStats) {
            body.appendParagraph(JSON.stringify(prevStats));
            totalRecords++;
          }
        }
        
      } catch (err) {
        Logger.log("  ✗ Ошибка: " + err.message);
      }
    }
    
    body.appendParagraph("");
    body.appendParagraph("=" .repeat(80));
    body.appendParagraph("Всего записей: " + totalRecords);
    
    var duration = Math.round((new Date().getTime() - startTime) / 1000);
    Logger.log("✅ Собрано записей: " + totalRecords + " за " + duration + " сек");
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "✅ Статистика собрана: " + totalRecords + " записей",
      "Готово",
      5
    );
    
  } catch (err) {
    Logger.log("❌ Ошибка: " + err);
    SpreadsheetApp.getActiveSpreadsheet().toast("❌ " + err.message, "Ошибка", 5);
  }
}

/**
 * Собирает данные с листа статистики в структурированном виде
 */
function collectStatsFromSheet(sheet, operatorName, period) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return null;
    
    var data = sheet.getRange(1, 1, Math.min(lastRow, 40), 2).getValues();
    
    var stats = {
      operator: operatorName,
      type: "summary",
      period: period,
      updated: "",
      totalFixes: 0,
      uniqueRecords: 0,
      open: 0,
      closed: 0,
      statuses: {}
    };
    
    for (var i = 0; i < data.length; i++) {
      var a = String(data[i][0] || "").toLowerCase().trim();
      var b = data[i][1];
      var bNum = parseInt(b) || 0;
      
      if (a.indexOf("обновлено") > -1) stats.updated = String(data[i][0]);
      if (a.indexOf("всего фиксаций") > -1) stats.totalFixes = bNum;
      if (a.indexOf("уникальных записей") > -1) stats.uniqueRecords = bNum;
      
      if (a.indexOf("↳") > -1 || a.indexOf("из них") > -1) {
        if (a.indexOf("открыто") > -1) stats.open = bNum;
        if (a.indexOf("закрыто") > -1) stats.closed = bNum;
      }
      
      if (a.indexOf("•") > -1) {
        if (a.indexOf("отрицательный") > -1) stats.statuses.negative = bNum;
        else if (a.indexOf("положительный") > -1) stats.statuses.positive = bNum;
        else if (a.indexOf("заявка закрыта") > -1) stats.statuses.closed = bNum;
        else if (a.indexOf("открыть карту") > -1) stats.statuses.openCard = bNum;
        else if (a.indexOf("тиббиёт") > -1) stats.statuses.medical = bNum;
      }
    }
    
    if (stats.totalFixes === 0) return null;
    
    return stats;
    
  } catch (err) {
    Logger.log("Ошибка сбора статистики: " + err);
    return null;
  }
}

/**
 * Очищает документ Google Docs
 */
function clearDocsDocument() {
  if (DOCS_ID === "ВСТАВЬТЕ_ID_ДОКУМЕНТА_СЮДА") {
    SpreadsheetApp.getUi().alert("⚠️ Необходимо настроить DOCS_ID");
    return;
  }
  
  var result = SpreadsheetApp.getUi().alert(
    "Очистка документа",
    "Вы уверены что хотите очистить документ?\nВсе данные будут удалены.",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (result === SpreadsheetApp.getUi().Button.YES) {
    try {
      var doc = DocumentApp.openById(DOCS_ID);
      doc.getBody().clear();
      
      Logger.log("✓ Документ очищен");
      SpreadsheetApp.getActiveSpreadsheet().toast("✅ Документ очищен", "Готово", 3);
      
    } catch (err) {
      Logger.log("❌ Ошибка очистки: " + err);
      SpreadsheetApp.getActiveSpreadsheet().toast("❌ " + err.message, "Ошибка", 5);
    }
  }
}

// =============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Получает список операторов из листа Настройки
 */
function getOperatorList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!sheet) {
    throw new Error("Лист 'Настройки' не найден");
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("В листе 'Настройки' нет данных");
  }
  
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var operators = [];
  
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0] || "").trim();
    var id = String(data[i][1] || "").trim();
    var status = String(data[i][2] || "активен").trim();
    
    if (name && id && id !== "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ") {
      operators.push({
        name: name,
        spreadsheetId: id,
        status: status
      });
    }
  }
  
  return operators;
}

/**
 * Форматирует дату в строку
 */
function formatDate(value) {
  if (!value) return null;
  
  var date = null;
  
  if (value instanceof Date) {
    date = value;
  } else {
    var str = String(value).trim();
    
    // Формат: "01.12.2024 10:30:45"
    var match1 = str.match(/(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/);
    if (match1) {
      date = new Date(
        parseInt(match1[3], 10),
        parseInt(match1[2], 10) - 1,
        parseInt(match1[1], 10),
        parseInt(match1[4], 10),
        parseInt(match1[5], 10),
        parseInt(match1[6], 10)
      );
    } else {
      // Формат: "01.12.2024"
      var match2 = str.match(/(\d{2})\.(\d{2})\.(\d{4})/);
      if (match2) {
        date = new Date(
          parseInt(match2[3], 10),
          parseInt(match2[2], 10) - 1,
          parseInt(match2[1], 10)
        );
      } else {
        date = new Date(str);
      }
    }
  }
  
  if (!date || isNaN(date.getTime()) || date.getFullYear() < 2000) {
    return null;
  }
  
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
}

/**
 * Показывает инструкцию по использованию
 */
function showInstructions() {
  var html = HtmlService.createHtmlOutput(
    '<h2>📄 Инструкция по использованию Docs Collector</h2>' +
    '<h3>Шаг 1: Создание документа</h3>' +
    '<ol>' +
    '<li>Откройте <a href="https://docs.google.com" target="_blank">Google Docs</a></li>' +
    '<li>Создайте новый пустой документ</li>' +
    '<li>Назовите его "Архив данных операторов"</li>' +
    '<li>Скопируйте ID из URL (часть между /d/ и /edit)</li>' +
    '</ol>' +
    '<h3>Шаг 2: Настройка скрипта</h3>' +
    '<ol>' +
    '<li>Откройте Расширения → Apps Script</li>' +
    '<li>Найдите файл docs_collector.gs</li>' +
    '<li>Вставьте ID документа в константу DOCS_ID</li>' +
    '<li>Сохраните (Ctrl+S)</li>' +
    '</ol>' +
    '<h3>Шаг 3: Сбор данных</h3>' +
    '<ol>' +
    '<li>Обновите страницу таблицы</li>' +
    '<li>Откройте меню "📄 Docs Collector"</li>' +
    '<li>Выберите "🔄 Собрать все данные в Docs"</li>' +
    '<li>Дождитесь завершения</li>' +
    '</ol>' +
    '<h3>Шаг 4: Работа с Python</h3>' +
    '<ol>' +
    '<li>Используйте Google Docs API для чтения данных</li>' +
    '<li>Каждая строка = JSON запись</li>' +
    '<li>Обрабатывайте данные в Python</li>' +
    '<li>Записывайте результаты обратно в Sheets</li>' +
    '</ol>' +
    '<p><strong>Формат данных:</strong></p>' +
    '<pre>{"operator":"Иванов","sheet":"11.2024","card":"1234","status":"положительный","date":"01.11.2024 10:30:00"}</pre>' +
    '<p><strong>Лимиты:</strong></p>' +
    '<ul>' +
    '<li>Максимум 10,000 записей за запуск</li>' +
    '<li>Время выполнения до 6 минут</li>' +
    '<li>Размер документа до 1 МБ</li>' +
    '</ul>'
  )
  .setWidth(600)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Инструкция');
}

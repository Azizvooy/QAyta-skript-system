/**
 * =============================================================================
 * ПОЛНЫЙ СКРИПТ ДЛЯ GOOGLE SHEETS - СИСТЕМА ФИКСАЦИИ
 * =============================================================================
 * Версия: 3.0
 * Дата: 24.11.2025
 * 
 * ФУНКЦИОНАЛ:
 * - Автоматическая фиксация при редактировании (onEdit)
 * - Заполнение пропущенных дат и ФИО (каждые 10 минут)
 * - Статистика текущего месяца (каждый час)
 * - Защита прошлых данных (ежедневно в 00:00)
 * - Автоматическое архивирование (19 числа в 23:00)
 * 
 * ОСОБЕННОСТИ ЗАЩИТЫ:
 * - Колонки H и I (ФИО и Дата) остаются редактируемыми для всех пользователей
 * - Остальные колонки защищены только для прошлых дней
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Sheets → Расширения → Apps Script
 * 2. Удалите весь код и вставьте этот скрипт
 * 3. Сохраните (Ctrl+S)
 * 4. Обновите страницу таблицы
 * 5. В меню появится "⚙️ Система Фиксаций"
 * 6. Запустите "⚙️ Настроить систему" ОДИН РАЗ
 * =============================================================================
 */

// =============================================================================
// ГЛОБАЛЬНЫЕ НАСТРОЙКИ
// =============================================================================

// ID таблицы для сбора статистики
var STATISTICS_SPREADSHEET_ID = "1wlqqSCV3HW5ZgfYUT6IS2Ne466jJQeEKH1Nl4Tx2jdc";

var CLOSED_STATUS_LIST = [
  "отрицательный",
  "положительный",
  "заявка закрыта (не удалось дозвониться)",
  "открыть карту",
  "тиббиёт ходими аризаси"
];

var CLOSED_STATUS_SET = new Set(CLOSED_STATUS_LIST.map(function (s) {
  return s.toLowerCase();
}));

var STATUS_KEYS = CLOSED_STATUS_LIST.slice();

var PROTECTION_DESCRIPTION = "FIKSA auto-lock before today";
var PROTECTION_ARCHIVE_DESCRIPTION = "Archive sheet lock";

var LOOKUP_FORMULA_C =
  '=IF(B2="";"";IFERROR(INDEX(\'Аризалар\'!B:B;MATCH(REGEXREPLACE(TRIM(SUBSTITUTE(SUBSTITUTE(TO_TEXT(B2);CHAR(160);" ");CHAR(8203);" ")); "\\s+"; " ");ARRAYFORMULA(REGEXREPLACE(TRIM(SUBSTITUTE(SUBSTITUTE(TO_TEXT(\'Аризалар\'!C:C);CHAR(160);" ");CHAR(8203);" ")); "\\s+"; " "));0));""))';

var LOOKUP_FORMULA_D =
  '=IF(B2="";"";IFERROR(INDEX(\'Аризалар\'!A:A;MATCH(REGEXREPLACE(TRIM(SUBSTITUTE(SUBSTITUTE(TO_TEXT(B2);CHAR(160);" ");CHAR(8203);" ")); "\\s+"; " ");ARRAYFORMULA(REGEXREPLACE(TRIM(SUBSTITUTE(SUBSTITUTE(TO_TEXT(\'Аризалар\'!C:C);CHAR(160);" ");CHAR(8203);" ")); "\\s+"; " "));0));""))';

// Названия месяцев на русском
var MONTH_NAMES = [
  "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
];

// =============================================================================
// ФУНКЦИЯ ИНИЦИАЛИЗАЦИИ
// =============================================================================

/**
 * Главная функция настройки - удаляет все старые триггеры и создаёт новые
 * ЗАПУСКАЕТСЯ АВТОМАТИЧЕСКИ при onInstall ИЛИ ВРУЧНУЮ из меню
 */
function setupAllTriggers() {
  // Удаляем ВСЕ существующие триггеры
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log("✓ Все старые триггеры удалены");

  // 1. Триггер заполнения пропущенных дат и ФИО (каждые 10 минут)
  ScriptApp.newTrigger("fillMissedDatesFast")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("✓ Триггер fillMissedDatesFast создан (каждые 10 минут)");

  // 2. Триггер обновления статистики текущего месяца (каждый час)
  ScriptApp.newTrigger("updateCurrentMonthStatistics")
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log("✓ Триггер updateCurrentMonthStatistics создан (каждый час)");

  // 3. Триггер защиты прошлых строк (ежедневно в 00:00)
  ScriptApp.newTrigger("protectPastRows")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
  Logger.log("✓ Триггер protectPastRows создан (ежедневно в 00:00)");

  // 4. Триггер архивирования (каждое 19 число месяца в 23:00)
  ScriptApp.newTrigger("transferFiksa")
    .timeBased()
    .onMonthDay(19)
    .atHour(23)
    .create();
  Logger.log("✓ Триггер transferFiksa создан (19 числа каждого месяца в 23:00)");

  // 5. Триггер заполнения архивных данных (ежедневно в 02:00)
  ScriptApp.newTrigger("fillArchiveDataAuto")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log("✓ Триггер fillArchiveDataAuto создан (ежедневно в 02:00)");

  // 6. Триггер создания сводки по дням (ежедневно в 03:00)
  ScriptApp.newTrigger("createDailySummarySheetAuto")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  Logger.log("✓ Триггер createDailySummarySheetAuto создан (ежедневно в 03:00)");

  Logger.log("\n=== НАСТРОЙКА ЗАВЕРШЕНА ===");
  
  // Применяем цветовое форматирование к колонке E
  applyStatusColorFormatting();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Все триггеры настроены!\n\n" +
    "• Заполнение пропусков - каждые 10 минут\n" +
    "• Статистика - каждый час\n" +
    "• Защита данных - ежедневно в 00:00\n" +
    "• Архивирование - 19 числа в 23:00\n" +
    "• Заполнение архивов - ежедневно в 02:00\n" +
    "• Сводка по дням - ежедневно в 03:00\n\n" +
    "⚠️ Функции запустятся автоматически по расписанию",
    "✅ Настройка завершена",
    15
  );
  
  Logger.log("Настройка завершена успешно. Функции запустятся по триггерам.");
}

// =============================================================================
// УСЛОВНОЕ ФОРМАТИРОВАНИЕ ДЛЯ КОЛОНКИ E (СТАТУСЫ)
// =============================================================================

/**
 * Применяет цветовое форматирование к колонке E с выпадающим списком статусов
 * Цвета для каждого статуса:
 * - Отрицательный: красный (#ff0000)
 * - Положительный: зеленый (#00ff00)
 * - Тишине: нежно-красный (#ffcccc)
 * - Соед прервано: нежно-красный (#ffcccc)
 * - НЕТ ОТВЕТА (ЗАНЯТО): желтый (#ffff00)
 * - Заявка закрыта: серый (#cccccc)
 * - Открыть карту: небесный (#87ceeb)
 * - Тиббиёт ходими аризаси: голубой нежный (#add8e6)
 */
function applyStatusColorFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FIKSA");
  
  if (!sheet) {
    Logger.log("Лист FIKSA не найден для применения форматирования");
    return;
  }
  
  var lastRow = sheet.getMaxRows();
  if (lastRow < 2) lastRow = 1000; // Минимум 1000 строк
  
  // Удаляем все существующие правила условного форматирования для колонки E
  var rules = sheet.getConditionalFormatRules();
  var newRules = [];
  
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var keepRule = true;
    
    for (var j = 0; j < ranges.length; j++) {
      if (ranges[j].getColumn() === 5) { // Колонка E
        keepRule = false;
        break;
      }
    }
    
    if (keepRule) {
      newRules.push(rules[i]);
    }
  }
  
  // Диапазон для колонки E (начиная со строки 2)
  var range = sheet.getRange("E2:E" + lastRow);
  
  // Создаем правила условного форматирования для каждого статуса
  
  // 1. Отрицательный - красный
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("отрицательный")
    .setBackground("#ff0000")
    .setFontColor("#ffffff")
    .setRanges([range])
    .build();
  newRules.push(rule1);
  
  // 2. Положительный - зеленый
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("положительный")
    .setBackground("#00ff00")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule2);
  
  // 3. Тишине - нежно-красный
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("тишине")
    .setBackground("#ffcccc")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule3);
  
  // 4. Соед прервано - нежно-красный
  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("соед прервано")
    .setBackground("#ffcccc")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule4);
  
  // 5. НЕТ ОТВЕТА (ЗАНЯТО) - желтый (выделяемый)
  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("НЕТ ОТВЕТА (ЗАНЯТО)")
    .setBackground("#ffff00")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule5);
  
  // 6. Заявка закрыта (не удалось дозвониться) - серый
  var rule6 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("заявка закрыта (не удалось дозвониться)")
    .setBackground("#cccccc")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule6);
  
  // 7. Открыть карту - небесный
  var rule7 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("открыть карту")
    .setBackground("#87ceeb")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule7);
  
  // 8. Тиббиёт ходими аризаси - голубой нежный
  var rule8 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("тиббиёт ходими аризаси")
    .setBackground("#add8e6")
    .setFontColor("#000000")
    .setRanges([range])
    .build();
  newRules.push(rule8);
  
  // Применяем все правила
  sheet.setConditionalFormatRules(newRules);
  
  Logger.log("✓ Цветовое форматирование применено к колонке E (8 статусов)");
}

// =============================================================================
// МЕНЮ
// =============================================================================

/**
 * Вызывается автоматически при установке надстройки или первом копировании таблицы
 * Настраивает все триггеры автоматически
 */
function onInstall(e) {
  onOpen(e);
  setupAllTriggers();
}

/**
 * Создаёт меню при открытии таблицы
 * ПОЛЬЗОВАТЕЛИ видят только обновление статистики
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("⚙️ Система Фиксаций")
    .addItem("📊 Обновить статистику", "updateCurrentMonthStatistics")
    .addSeparator()
    .addItem("🎨 Применить цвета к статусам", "applyStatusColorFormatting")
    .addToUi();
}

// =============================================================================
// 1. ONEDIT - АВТОМАТИЧЕСКАЯ ФИКСАЦИЯ ПРИ РЕДАКТИРОВАНИИ
// =============================================================================

/**
 * Триггер onEdit - автоматическая фиксация при редактировании
 * Заполняет:
 * - Колонку I (дата/время) при заполнении B (номер карты)
 * - Колонку H (ФИО оператора) при заполнении E (статус) закрытым статусом
 */
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() !== "FIKSA") return;

    var editedCol = range.getColumn();
    var numCols = range.getNumColumns();
    var startRow = range.getRow();
    var numRows = range.getNumRows();

    var touchesB = (editedCol <= 2 && (editedCol + numCols - 1) >= 2);
    var touchesE = (editedCol <= 5 && (editedCol + numCols - 1) >= 5);

    if (!touchesB && !touchesE) return;

    var ss = e.source;
    var settingSheet = ss.getSheetByName("SETTING");
    if (!settingSheet) return;
    var settingValue = settingSheet.getRange("B2").getValue();

    var allowed = CLOSED_STATUS_LIST.map(function(s){ return s.toLowerCase(); });
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");

    var bValues = touchesB ? sheet.getRange(startRow, 2, numRows, 1).getValues() : null;
    var eValues = touchesE ? sheet.getRange(startRow, 5, numRows, 1).getValues() : null;
    var iCurrent = sheet.getRange(startRow, 9, numRows, 1).getValues();
    var hCurrent = sheet.getRange(startRow, 8, numRows, 1).getValues();

    var hOut = [];
    var iOut = [];

    for (var i = 0; i < numRows; i++) {
      iOut[i] = [ iCurrent[i][0] ];
      hOut[i] = [ hCurrent[i][0] ];
    }

    for (var i = 0; i < numRows; i++) {
      if (touchesB) {
        var bVal = bValues[i][0];
        var bNotEmpty = String(bVal).toString().trim() !== "";
        var iValNow = iCurrent[i][0];

        if (bNotEmpty) {
          if (iValNow === "" || iValNow === null || typeof iValNow === "undefined") {
            iOut[i][0] = timestamp;
          } else {
            iOut[i][0] = iValNow;
          }
        } else {
          iOut[i][0] = "";
        }
      }

      if (touchesE) {
        var eVal = String(eValues[i][0]).trim();
        if (eVal === "") {
          hOut[i][0] = "";
        } else {
          var isAllowed = allowed.indexOf(eVal.toLowerCase()) !== -1;
          if (isAllowed) {
            hOut[i][0] = settingValue;
          } else {
            hOut[i][0] = "";
          }
        }
      }
    }

    if (touchesB) {
      sheet.getRange(startRow, 9, numRows, 1).setValues(iOut);
    }

    if (touchesE) {
      sheet.getRange(startRow, 8, numRows, 1).setValues(hOut);
    }

  } catch (err) {
    console.error("onEdit error: " + err);
  }
}

// =============================================================================
// 2. ЗАПОЛНЕНИЕ ПРОПУЩЕННЫХ ДАТ И ФИО
// =============================================================================

/**
 * Заполняет пропущенные даты в колонке I
 * Заполняет пропущенные/ошибочные ФИО в колонке H
 * Улучшенная версия с проверкой ошибок
 */
function fillMissedDatesFast() {
  var startTime = new Date().getTime();
  
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("FIKSA");
  if (!sheet) {
    Logger.log("Лист FIKSA не найден");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Нет данных для обработки");
    return;
  }

  const settingSheet = ss.getSheetByName("SETTING");
  if (!settingSheet) {
    Logger.log("Лист SETTING не найден");
    return;
  }
  const settingValue = settingSheet.getRange("B2").getValue();

  const allowed = CLOSED_STATUS_LIST.map(s => s.toLowerCase());

  // Загружаем только нужные столбцы одним запросом
  let allData = sheet.getRange(2, 2, lastRow - 1, 8).getValues(); // B-I (столбцы 2-9)

  let previousDate = null;
  let changedI = false;
  let changedH = false;

  for (let r = 0; r < allData.length; r++) {
    let colB = allData[r][0];  // B
    let colE = allData[r][3];  // E
    let colH = allData[r][6];  // H
    let colI = allData[r][7];  // I

    // Обработка колонки I (дата/время)
    if (colB) {
      if (colI) {
        previousDate = colI;
      } else if (previousDate) {
        allData[r][7] = previousDate;
        changedI = true;
      }
    }

    // Обработка колонки H (ФИО)
    const eVal = String(colE).trim();
    
    // Проверяем, является ли значение в H ошибкой или пустым
    const hVal = String(colH).trim();
    const isError = hVal.indexOf("#") === 0 || hVal.toLowerCase().indexOf("error") !== -1 || hVal.toLowerCase().indexOf("ошибка") !== -1;
    const isEmpty = !colH || hVal === "";
    
    if (eVal !== "") {
      const isAllowed = allowed.indexOf(eVal.toLowerCase()) !== -1;
      if (isAllowed && (isEmpty || isError)) {
        allData[r][6] = settingValue;
        changedH = true;
      }
    }
  }

  // Записываем только если были изменения
  if (changedI || changedH) {
    let updateData = allData.map(row => [row[6], row[7]]);
    sheet.getRange(2, 8, updateData.length, 2).setValues(updateData);
    
    var endTime = new Date().getTime();
    Logger.log("fillMissedDatesFast выполнен за " + (endTime - startTime) + " мс");
    
    ss.toast(
      "Заполнено пропусков за " + (endTime - startTime) + " мс",
      "✅ Готово",
      3
    );
  } else {
    var endTime = new Date().getTime();
    Logger.log("fillMissedDatesFast: изменений не требуется. Проверка за " + (endTime - startTime) + " мс");
  }
}

// =============================================================================
// 3. СТАТИСТИКА ТЕКУЩЕГО МЕСЯЦА
// =============================================================================

/**
 * Обновляет статистику текущего месяца (FIKSA)
 */
function updateCurrentMonthStatistics() {
  var startTime = new Date().getTime();
  
  try {
    // СНАЧАЛА ЗАПОЛНЯЕМ ПРОПУСКИ В ДАННЫХ
    Logger.log("updateCurrentMonthStatistics: Запуск заполнения пропусков...");
    fillMissedDatesFast();
    Logger.log("updateCurrentMonthStatistics: Заполнение пропусков завершено");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var fiksaSheet = ss.getSheetByName("FIKSA");
    
    if (!fiksaSheet) {
      Logger.log("Лист FIKSA не найден");
      return;
    }

    var period = getPeriodBounds(new Date(), 0);
    var monthName = getMonthNameFromPeriod(period.endDate);
    
    // Используем стандартное название для совместимости со скриптом сбора
    var statsSheetName = "Статистика";
    
    var statsSheet = ss.getSheetByName(statsSheetName);
    if (!statsSheet) {
      statsSheet = ss.insertSheet(statsSheetName);
      // Перемещаем лист в начало
      ss.setActiveSheet(statsSheet);
      ss.moveActiveSheet(1);
    }

    var data = readSheetData(fiksaSheet);
    if (!data) {
      Logger.log("Нет данных для анализа");
      return;
    }

    var statsData = analyzeDataOptimized(data, period.startDate, period.endDate);
    displayStatisticsOptimized(
      statsSheet,
      statsData,
      period.startDate,
      period.endDate,
      "СТАТИСТИКА ФИКСАЦИИ - " + monthName.toUpperCase()
    );

    // УДАЛЯЕМ ЛИШНИЕ ПУСТЫЕ СТРОКИ И СТОЛБЦЫ
    cleanupSheet(statsSheet);

    // СОЗДАЕМ/ОБНОВЛЯЕМ ЛИСТ ПРЕДЫДУЩЕГО МЕСЯЦА
    updatePreviousMonthStatistics(ss);

    // УПРАВЛЯЕМ ВИДИМОСТЬЮ ЛИСТОВ
    Logger.log("updateCurrentMonthStatistics: Запуск скрытия листов...");
    manageSheetVisibility(ss);
    Logger.log("updateCurrentMonthStatistics: Скрытие листов завершено");

    var endTime = new Date().getTime();
    Logger.log("updateCurrentMonthStatistics выполнен за " + (endTime - startTime) + " мс");
    
    ss.toast(
      "Статистика за " + monthName + " обновлена за " + Math.round((endTime - startTime) / 1000) + " сек",
      "✅ Текущий месяц",
      5
    );
  } catch (err) {
    Logger.log("Ошибка в updateCurrentMonthStatistics: " + err);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Ошибка: " + err,
      "❌ Ошибка",
      5
    );
  }
}

/**
 * Получает название месяца из периода
 */
function getMonthNameFromPeriod(endDate) {
  var month = endDate.getMonth();
  var year = endDate.getFullYear();
  return MONTH_NAMES[month] + " " + year;
}

/**
 * Создает/обновляет лист "Предыдущий месяц" для совместимости со скриптом сбора
 */
function updatePreviousMonthStatistics(ss) {
  try {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    var fiksaSheet = ss.getSheetByName("FIKSA");
    
    if (!fiksaSheet) {
      Logger.log("updatePreviousMonthStatistics: Лист FIKSA не найден");
      return;
    }

    var prevPeriod = getPeriodBounds(new Date(), -1);
    var prevStatsSheetName = "Предыдущий месяц";
    
    var prevStatsSheet = ss.getSheetByName(prevStatsSheetName);
    if (!prevStatsSheet) {
      prevStatsSheet = ss.insertSheet(prevStatsSheetName);
    }

    // Проверяем наличие архивного листа
    var settingSheet = ss.getSheetByName("SETTING");
    if (!settingSheet) {
      Logger.log("updatePreviousMonthStatistics: Лист SETTING не найден");
      return;
    }

    var baseName = (settingSheet.getRange("B2").getValue() || "").toString().trim();
    if (!baseName) {
      Logger.log("updatePreviousMonthStatistics: SETTING!B2 пуст");
      return;
    }

    var archiveSheetName = getArchiveSheetName(baseName, prevPeriod.endDate);
    var archiveSheet = ss.getSheetByName(archiveSheetName);

    if (!archiveSheet) {
      Logger.log("updatePreviousMonthStatistics: Архивный лист '" + archiveSheetName + "' не найден");
      return;
    }

    var data = readSheetData(archiveSheet);
    if (!data) {
      Logger.log("updatePreviousMonthStatistics: Архив пуст");
      return;
    }

    var statsData = analyzeDataOptimized(data, prevPeriod.startDate, prevPeriod.endDate);
    displayStatisticsOptimized(
      prevStatsSheet,
      statsData,
      prevPeriod.startDate,
      prevPeriod.endDate,
      "СТАТИСТИКА ФИКСАЦИИ - " + getMonthNameFromPeriod(prevPeriod.endDate).toUpperCase()
    );

    cleanupSheet(prevStatsSheet);
    
    Logger.log("updatePreviousMonthStatistics: Статистика предыдущего месяца обновлена");
  } catch (err) {
    Logger.log("Ошибка в updatePreviousMonthStatistics: " + err);
  }
}

/**
 * Получает имя архивного листа
 */
function getArchiveSheetName(baseName, endDate) {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  return sanitizeSheetName(baseName + " " + Utilities.formatDate(endDate, tz, "MM.yyyy"));
}

/**
 * Автоматическое создание сводки по дням (запускается по триггеру)
 */
function createDailySummarySheetAuto() {
  var startTime = new Date().getTime();
  
  try {
    createDailySummarySheet();
    
    var endTime = new Date().getTime();
    Logger.log("createDailySummarySheetAuto завершено за " + 
               Math.round((endTime - startTime) / 1000) + " сек");
  } catch (err) {
    Logger.log("Ошибка в createDailySummarySheetAuto: " + err);
  }
}

/**
 * Создает сводный лист со статистикой по дням для всех операторов
 */
function createDailySummarySheet(ss) {
  try {
    // СНАЧАЛА ЗАПОЛНЯЕМ ПРОПУСКИ В ДАННЫХ
    Logger.log("createDailySummarySheet: Запуск заполнения пропусков...");
    fillMissedDatesFast();
    Logger.log("createDailySummarySheet: Заполнение пропусков завершено");
    
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = Session.getScriptTimeZone();
    
    var summarySheetName = "Сводка по дням";
    var summarySheet = ss.getSheetByName(summarySheetName);
    
    if (!summarySheet) {
      summarySheet = ss.insertSheet(summarySheetName);
    }
    
    summarySheet.clear();
    
    // Получаем текущий и предыдущий периоды
    var currentPeriod = getPeriodBounds(new Date(), 0);
    var prevPeriod = getPeriodBounds(new Date(), -1);
    
    // Собираем данные со всех листов операторов
    var allData = [];
    var sheets = ss.getSheets();
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      
      // Пропускаем служебные листы
      if (sheetName === "Аризалар" || 
          sheetName === "SETTING" || 
          sheetName === "Статистика" || 
          sheetName === "Предыдущий месяц" ||
          sheetName === summarySheetName) {
        continue;
      }
      
      // Определяем оператора и период из имени листа
      var operatorName = "";
      var period = null;
      
      // Проверяем, это архивный лист с датой в названии
      if (/\d{2}\.\d{4}/.test(sheetName)) {
        // Архивный лист - извлекаем имя оператора
        operatorName = sheetName.replace(/\s+\d{2}\.\d{4}$/, '').trim();
        
        // Читаем данные с архивного листа
        var data = readSheetData(sheet);
        if (data && data.length > 0) {
          // ГИБКИЙ АНАЛИЗ: определяем реальный период данных из самих данных
          var minDate = null;
          var maxDate = null;
          
          for (var d = 0; d < data.length; d++) {
            var recordDate = parseDateTime(data[d][8]);
            if (recordDate) {
              if (!minDate || recordDate < minDate) minDate = recordDate;
              if (!maxDate || recordDate > maxDate) maxDate = recordDate;
            }
          }
          
          if (minDate && maxDate) {
            // Расширяем период на несколько дней для захвата всех данных
            period = {
              startDate: new Date(minDate.getFullYear(), minDate.getMonth(), minDate.getDate() - 5),
              endDate: new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate() + 5)
            };
            
            // Анализируем данные по дням с реальным периодом
            var dayStats = analyzeDailyData(data, period.startDate, period.endDate, operatorName);
            allData = allData.concat(dayStats);
            Logger.log("Обработан архивный лист: " + sheetName + ", записей: " + dayStats.length + 
                      " (период: " + Utilities.formatDate(minDate, timeZone, "dd.MM.yyyy") + 
                      " - " + Utilities.formatDate(maxDate, timeZone, "dd.MM.yyyy") + ")");
          }
        }
      }
    }
    
    // Добавляем данные из текущего листа FIKSA
    var fiksaSheet = ss.getSheetByName("FIKSA");
    if (fiksaSheet) {
      var settingSheet = ss.getSheetByName("SETTING");
      if (settingSheet) {
        var operatorName = (settingSheet.getRange("B2").getValue() || "").toString().trim();
        if (operatorName) {
          var fiksaData = readSheetData(fiksaSheet);
          if (fiksaData) {
            var fiksaDayStats = analyzeDailyData(fiksaData, currentPeriod.startDate, currentPeriod.endDate, operatorName);
            allData = allData.concat(fiksaDayStats);
          }
        }
      }
    }
    
    // Сортируем по дате (по убыванию - новые сверху)
    allData.sort(function(a, b) {
      return b.date - a.date;
    });
    
    // Формируем заголовки
    var headers = [
      "Дата",
      "ФИО",
      "Всего фиксаций",
      "Уникальных карт",
      "Закрытых",
      "Открытых",
      "Повторных"
    ];
    
    var rows = [headers];
    
    // Добавляем данные
    for (var i = 0; i < allData.length; i++) {
      var item = allData[i];
      rows.push([
        Utilities.formatDate(item.date, timeZone, "dd.MM.yyyy (EEE)"),
        item.operator,
        item.totalFixes,
        item.uniqueCards,
        item.closedCards,
        item.openCards,
        item.repeatCards
      ]);
    }
    
    // Записываем данные
    if (rows.length > 0) {
      summarySheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
      
      // Форматирование
      summarySheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#4a86e8")
        .setFontColor("#ffffff");
      
      summarySheet.setFrozenRows(1);
      summarySheet.autoResizeColumns(1, headers.length);
    }
    
    cleanupSheet(summarySheet);
    
    Logger.log("createDailySummarySheet: Сводка создана, записей: " + (rows.length - 1));
    
  } catch (err) {
    Logger.log("Ошибка в createDailySummarySheet: " + err);
  }
}

/**
 * Анализирует данные по дням для одного оператора
 */
function analyzeDailyData(data, startDate, endDate, operatorName) {
  var timeZone = Session.getScriptTimeZone();
  var dayStats = {};
  var allCards = {}; // Все карты с их первой датой
  
  // Первый проход - собираем все карты и их первые даты
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var bRaw = row[1];
    if (!bRaw) continue;
    
    var bValue = String(bRaw).trim();
    if (!bValue) continue;
    
    var recordDate = parseDateTime(row[8]);
    if (!recordDate || recordDate < startDate || recordDate > endDate) continue;
    
    if (!allCards[bValue] || recordDate < allCards[bValue]) {
      allCards[bValue] = recordDate;
    }
  }
  
  // Второй проход - собираем статистику по дням
  for (var j = 0; j < data.length; j++) {
    var record = data[j];
    var bCell = record[1];
    if (!bCell) continue;
    
    var bStr = String(bCell).trim();
    if (!bStr) continue;
    
    var dateObj = parseDateTime(record[8]);
    if (!dateObj || dateObj < startDate || dateObj > endDate) continue;
    
    var dayKey = Utilities.formatDate(dateObj, timeZone, "yyyy-MM-dd");
    var status = String(record[4] || "").trim().toLowerCase();
    
    if (!dayStats[dayKey]) {
      dayStats[dayKey] = {
        date: dateObj,
        operator: operatorName,
        totalFixes: 0,
        uniqueCards: {},
        closedCards: 0,
        openCards: 0,
        repeatCards: 0,
        closedSet: {},
        openSet: {}
      };
    }
    
    var dayStat = dayStats[dayKey];
    dayStat.totalFixes++;
    dayStat.uniqueCards[bStr] = true;
    
    // Проверяем, была ли эта карта в другой день
    var firstDate = allCards[bStr];
    if (firstDate) {
      var firstDayKey = Utilities.formatDate(firstDate, timeZone, "yyyy-MM-dd");
      if (firstDayKey !== dayKey) {
        dayStat.repeatCards++;
      }
    }
    
    // Подсчитываем закрытые/открытые
    if (CLOSED_STATUS_SET.has(status)) {
      dayStat.closedSet[bStr] = true;
    }
  }
  
  // Финализируем статистику
  var result = [];
  for (var dayKey in dayStats) {
    var stat = dayStats[dayKey];
    var uniqueCount = Object.keys(stat.uniqueCards).length;
    var closedCount = Object.keys(stat.closedSet).length;
    
    result.push({
      date: stat.date,
      operator: stat.operator,
      totalFixes: stat.totalFixes,
      uniqueCards: uniqueCount,
      closedCards: closedCount,
      openCards: uniqueCount - closedCount,
      repeatCards: stat.repeatCards
    });
  }
  
  return result;
}

// Вспомогательные функции для статистики
function getPeriodBounds(referenceDate, offset) {
  var date = new Date(referenceDate);
  var year = date.getFullYear();
  var month = date.getMonth();
  if (date.getDate() < 20) {
    month -= 1;
    if (month < 0) {
      month = 11;
      year -= 1;
    }
  }

  month += offset;
  while (month < 0) {
    month += 12;
    year -= 1;
  }
  while (month > 11) {
    month -= 12;
    year += 1;
  }

  var startDate = new Date(year, month, 20, 0, 0, 0);
  var endMonth = month + 1;
  var endYear = year;
  if (endMonth > 11) {
    endMonth -= 12;
    endYear += 1;
  }
  var endDate = new Date(endYear, endMonth, 19, 23, 59, 59);

  return { startDate: startDate, endDate: endDate };
}

function readSheetData(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var columnB = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var filledRows = 0;
  for (var i = 0; i < columnB.length; i++) {
    if (columnB[i][0] && String(columnB[i][0]).trim() !== "") {
      filledRows = i + 1;
    }
  }
  if (filledRows === 0) return null;

  return sheet.getRange(2, 1, filledRows, 9).getValues();
}

/**
 * ОПТИМИЗИРОВАННАЯ версия анализа данных с отслеживанием первой фиксации
 */
function analyzeDataOptimized(data, startDate, endDate) {
  var timeZone = Session.getScriptTimeZone();

  var stats = {
    total: {
      totalFixes: 0,
      allUniqueB: {},
      closedB: {},
      openB: {},
      uniqueBWithStatus: {},
      firstFixDateByB: {}
    },
    byDay: {},
    employees: {},
    employeesByDay: {}
  };

  for (var s = 0; s < STATUS_KEYS.length; s++) {
    stats.total.uniqueBWithStatus[STATUS_KEYS[s]] = {};
  }

  // ПЕРВЫЙ ПРОХОД: определяем уникальные записи и их первые фиксации
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var bRaw = row[1];
    if (!bRaw) continue;

    var bValue = String(bRaw).trim();
    if (!bValue) continue;

    var recordDate = parseDateTime(row[8]);
    if (!recordDate || recordDate < startDate || recordDate > endDate) continue;

    stats.total.allUniqueB[bValue] = true;

    // Запоминаем дату первой фиксации
    if (!stats.total.firstFixDateByB[bValue]) {
      stats.total.firstFixDateByB[bValue] = recordDate;
    } else if (recordDate < stats.total.firstFixDateByB[bValue]) {
      stats.total.firstFixDateByB[bValue] = recordDate;
    }

    var status = String(row[4] || "").trim().toLowerCase();
    
    if (CLOSED_STATUS_SET.has(status)) {
      stats.total.closedB[bValue] = true;
      
      for (var s = 0; s < STATUS_KEYS.length; s++) {
        var statusKey = STATUS_KEYS[s];
        if (status === statusKey.toLowerCase()) {
          stats.total.uniqueBWithStatus[statusKey][bValue] = true;
          break;
        }
      }
    }
  }

  // Определяем открытые заявки
  for (var bKey in stats.total.allUniqueB) {
    if (!stats.total.closedB[bKey]) {
      stats.total.openB[bKey] = true;
    }
  }

  // ВТОРОЙ ПРОХОД: собираем детальную статистику
  for (var j = 0; j < data.length; j++) {
    var record = data[j];
    var bCell = record[1];
    if (!bCell) continue;

    var bStr = String(bCell).trim();
    if (!bStr) continue;

    var dateObj = parseDateTime(record[8]);
    if (!dateObj || dateObj < startDate || dateObj > endDate) continue;

    var dayKey = Utilities.formatDate(dateObj, timeZone, "yyyy-MM-dd");
    var employee = String(record[7] || "").trim();
    var statusValue = String(record[4] || "").trim().toLowerCase();

    stats.total.totalFixes++;

    if (!stats.byDay[dayKey]) {
      stats.byDay[dayKey] = {
        date: dateObj,
        totalFixes: 0,
        allUniqueB: {},
        closedB: {},
        openB: {},
        oldClosedCount: 0,
        oldClosedList: [],
        uniqueBWithStatus: {}
      };
      
      for (var s = 0; s < STATUS_KEYS.length; s++) {
        stats.byDay[dayKey].uniqueBWithStatus[STATUS_KEYS[s]] = {};
      }
    }

    var dayData = stats.byDay[dayKey];
    dayData.totalFixes++;
    dayData.allUniqueB[bStr] = true;

    if (CLOSED_STATUS_SET.has(statusValue)) {
      dayData.closedB[bStr] = true;
      
      // Проверяем, была ли первая фиксация в другой день
      var firstFixDate = stats.total.firstFixDateByB[bStr];
      if (firstFixDate) {
        var firstFixDayKey = Utilities.formatDate(firstFixDate, timeZone, "yyyy-MM-dd");
        
        if (firstFixDayKey !== dayKey) {
          if (dayData.oldClosedList.indexOf(bStr) === -1) {
            dayData.oldClosedCount++;
            dayData.oldClosedList.push(bStr);
          }
        }
      }
      
      for (var s = 0; s < STATUS_KEYS.length; s++) {
        var statusKey = STATUS_KEYS[s];
        if (statusValue === statusKey.toLowerCase()) {
          dayData.uniqueBWithStatus[statusKey][bStr] = true;
          break;
        }
      }
    }

    if (employee) {
      if (!stats.employees[employee]) {
        stats.employees[employee] = { timestamps: [], totalFixes: 0, uniqueCards: {} };
      }
      stats.employees[employee].timestamps.push(dateObj);
      stats.employees[employee].totalFixes++;
      stats.employees[employee].uniqueCards[bStr] = true;

      if (!stats.employeesByDay[employee]) {
        stats.employeesByDay[employee] = {};
      }
      if (!stats.employeesByDay[employee][dayKey]) {
        stats.employeesByDay[employee][dayKey] = { timestamps: [], totalFixes: 0, uniqueCards: {} };
      }
      stats.employeesByDay[employee][dayKey].timestamps.push(dateObj);
      stats.employeesByDay[employee][dayKey].totalFixes++;
      stats.employeesByDay[employee][dayKey].uniqueCards[bStr] = true;
    }
  }

  // Определяем открытые заявки для каждого дня
  for (var dayKey in stats.byDay) {
    var dayInfo = stats.byDay[dayKey];
    for (var bValue in dayInfo.allUniqueB) {
      if (!stats.total.closedB[bValue]) {
        dayInfo.openB[bValue] = true;
      }
    }
  }

  return stats;
}

/**
 * ОПТИМИЗИРОВАННАЯ версия отображения статистики с комментариями для открытых/закрытых
 */
function displayStatisticsOptimized(sheet, stats, startDate, endDate, title) {
  title = title || "СТАТИСТИКА ФИКСАЦИИ";
  sheet.clear();
  var timeZone = Session.getScriptTimeZone();

  var allRows = [];
  
  allRows.push([title]);
  allRows.push([
    "Период: " +
    Utilities.formatDate(startDate, timeZone, "dd.MM.yyyy") +
    " - " +
    Utilities.formatDate(endDate, timeZone, "dd.MM.yyyy")
  ]);
  allRows.push([
    "Обновлено: " +
    Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm:ss")
  ]);
  allRows.push([""]);

  allRows.push(["ОБЩАЯ СТАТИСТИКА"]);
  allRows.push(["Всего фиксаций:", stats.total.totalFixes]);
  allRows.push(["Всего уникальных записей (B):", Object.keys(stats.total.allUniqueB).length]);
  allRows.push(["  ↳ из них ОТКРЫТО (не закрыто):", Object.keys(stats.total.openB).length]);
  allRows.push(["  ↳ из них ЗАКРЫТО:", Object.keys(stats.total.closedB).length]);
  allRows.push(["  ↳ проверка (открыто + закрыто):", 
    Object.keys(stats.total.openB).length + Object.keys(stats.total.closedB).length]);
  allRows.push([""]);

  allRows.push(["УНИКАЛЬНЫЕ ПО СТАТУСАМ:"]);
  for (var s = 0; s < STATUS_KEYS.length; s++) {
    var key = STATUS_KEYS[s];
    allRows.push(["  • " + key + ":", Object.keys(stats.total.uniqueBWithStatus[key]).length]);
  }
  allRows.push([""]);
  allRows.push([""]);

  allRows.push(["СТАТИСТИКА ПО ДНЯМ"]);
  allRows.push([
    "Дата",
    "Всего фиксаций",
    "Уникальных всего",
    "Открыто",
    "Закрыто",
    "Отрицательный",
    "Положительный",
    "Заявка закрыта",
    "Открыть карту",
    "Тиббиёт ходими"
  ]);

  var dayStartRow = allRows.length + 1; // Запоминаем с какой строки начинаются дни

  var dayKeys = Object.keys(stats.byDay).sort();
  for (var d = 0; d < dayKeys.length; d++) {
    var dayKey = dayKeys[d];
    var dayData = stats.byDay[dayKey];
    
    var oldClosed = dayData.oldClosedCount || 0;
    var closedDisplay = Object.keys(dayData.closedB).length;
    if (oldClosed > 0) {
      closedDisplay = closedDisplay + " (" + oldClosed + ")";
    }
    
    allRows.push([
      Utilities.formatDate(dayData.date, timeZone, "dd.MM.yyyy (EEE)"),
      dayData.totalFixes,
      Object.keys(dayData.allUniqueB).length,
      Object.keys(dayData.openB).length,
      closedDisplay,
      Object.keys(dayData.uniqueBWithStatus["отрицательный"]).length,
      Object.keys(dayData.uniqueBWithStatus["положительный"]).length,
      Object.keys(dayData.uniqueBWithStatus["заявка закрыта (не удалось дозвониться)"]).length,
      Object.keys(dayData.uniqueBWithStatus["открыть карту"]).length,
      Object.keys(dayData.uniqueBWithStatus["тиббиёт ходими аризаси"]).length
    ]);
  }

  allRows.push([""]);
  allRows.push([""]);

  allRows.push(["СТАТИСТИКА ПО СОТРУДНИКАМ"]);
  allRows.push([
    "ФИО",
    "Всего фиксаций",
    "Уникальных карт"
  ]);

  var employeeNames = Object.keys(stats.employees).sort();
  for (var e = 0; e < employeeNames.length; e++) {
    var empName = employeeNames[e];
    var empData = stats.employees[empName];
    allRows.push([
      empName,
      empData.totalFixes,
      Object.keys(empData.uniqueCards).length
    ]);
  }

  allRows.push([""]);
  allRows.push([""]);

  allRows.push(["ДЕТАЛЬНАЯ СТАТИСТИКА ПО СОТРУДНИКАМ ПО ДНЯМ"]);

  for (var idx = 0; idx < employeeNames.length; idx++) {
    var name = employeeNames[idx];
    var dayDataByEmp = stats.employeesByDay[name];
    if (!dayDataByEmp) continue;

    allRows.push([name]);
    allRows.push([
      "Дата",
      "Фиксаций",
      "Уникальных карт"
    ]);

    var empDayKeys = Object.keys(dayDataByEmp).sort();
    for (var k = 0; k < empDayKeys.length; k++) {
      var empDayKey = empDayKeys[k];
      var dayInfo = dayDataByEmp[empDayKey];

      allRows.push([
        Utilities.formatDate(dayInfo.timestamps[0], timeZone, "dd.MM.yyyy (EEE)"),
        dayInfo.totalFixes,
        Object.keys(dayInfo.uniqueCards).length
      ]);
    }

    allRows.push([""]);
  }

  if (allRows.length > 0) {
    var maxCols = Math.max.apply(null, allRows.map(function(row) { return row.length; }));
    
    for (var i = 0; i < allRows.length; i++) {
      while (allRows[i].length < maxCols) {
        allRows[i].push("");
      }
    }
    
    sheet.getRange(1, 1, allRows.length, maxCols).setValues(allRows);
  }

  // Форматирование заголовков
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A5").setFontWeight("bold").setFontSize(12)
    .setBackground("#4a86e8").setFontColor("#ffffff");
  
  // ДОБАВЛЯЕМ КОММЕНТАРИИ К ЯЧЕЙКАМ С ОТКРЫТЫМИ/ЗАКРЫТЫМИ ЗАЯВКАМИ
  
  // Комментарий к общему количеству открытых
  var openBList = Object.keys(stats.total.openB).sort();
  if (openBList.length > 0) {
    sheet.getRange(8, 2).setNote("Открытые заявки:\n" + openBList.join(", "));
  }
  
  // Комментарий к общему количеству закрытых
  var closedBList = Object.keys(stats.total.closedB).sort();
  if (closedBList.length > 0) {
    sheet.getRange(9, 2).setNote("Закрытые заявки:\n" + closedBList.join(", "));
  }
  
  // Комментарии к открытым заявкам по дням
  for (var d = 0; d < dayKeys.length; d++) {
    var dayKey = dayKeys[d];
    var dayData = stats.byDay[dayKey];
    var rowNum = dayStartRow + d;
    
    // Комментарий к колонке "Открыто"
    var openList = Object.keys(dayData.openB).sort();
    if (openList.length > 0) {
      sheet.getRange(rowNum, 4).setNote("Открытые заявки:\n" + openList.join(", "));
    }
    
    // Комментарий к колонке "Закрыто" (включая старые)
    var closedList = Object.keys(dayData.closedB).sort();
    if (closedList.length > 0) {
      var noteText = "Закрытые заявки:\n" + closedList.join(", ");
      if (dayData.oldClosedCount > 0) {
        noteText += "\n\n⚠️ Закрыты старые заявки (" + dayData.oldClosedCount + "):\n" + 
                    dayData.oldClosedList.sort().join(", ");
      }
      sheet.getRange(rowNum, 5).setNote(noteText);
    }
  }
  
  sheet.autoResizeColumns(1, maxCols);
  sheet.setFrozenRows(4);
}

/**
 * Удаляет лишние пустые строки и столбцы из листа статистики
 */
function cleanupSheet(sheet) {
  try {
    // Находим последнюю заполненную строку
    var lastRow = sheet.getLastRow();
    var maxRows = sheet.getMaxRows();
    
    // Удаляем лишние строки (оставляем +10 запасных)
    if (maxRows > lastRow + 10) {
      sheet.deleteRows(lastRow + 11, maxRows - lastRow - 10);
      Logger.log("Удалено " + (maxRows - lastRow - 10) + " пустых строк из " + sheet.getName());
    }
    
    // Находим последний заполненный столбец
    var lastCol = sheet.getLastColumn();
    var maxCols = sheet.getMaxColumns();
    
    // Удаляем лишние столбцы (оставляем +5 запасных)
    if (maxCols > lastCol + 5) {
      sheet.deleteColumns(lastCol + 6, maxCols - lastCol - 5);
      Logger.log("Удалено " + (maxCols - lastCol - 5) + " пустых столбцов из " + sheet.getName());
    }
    
  } catch (e) {
    Logger.log("Ошибка при очистке листа " + sheet.getName() + ": " + e);
  }
}

function parseDateTime(value) {
  if (!value) return null;
  if (value instanceof Date) return value;

  var str = String(value).trim();
  var pattern = /(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/;
  var match = str.match(pattern);
  if (match) {
    var day = parseInt(match[1], 10);
    var month = parseInt(match[2], 10) - 1;
    var year = parseInt(match[3], 10);
    var hour = parseInt(match[4], 10);
    var minute = parseInt(match[5], 10);
    var second = parseInt(match[6], 10);
    return new Date(year, month, day, hour, minute, second);
  }
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function calculateOverallWorkTime(timestamps) {
  if (!timestamps || !timestamps.length) {
    return {
      totalWork: "0ч 0м",
      totalBreak: "0ч 0м",
      avgDayDuration: "0ч 0м"
    };
  }

  var timeZone = Session.getScriptTimeZone();
  var dayGroups = {};
  for (var i = 0; i < timestamps.length; i++) {
    var dayKey = Utilities.formatDate(timestamps[i], timeZone, "yyyy-MM-dd");
    if (!dayGroups[dayKey]) {
      dayGroups[dayKey] = [];
    }
    dayGroups[dayKey].push(timestamps[i]);
  }

  var totalWorkMinutes = 0;
  var totalBreakMinutes = 0;
  var dayKeys = Object.keys(dayGroups);

  for (var j = 0; j < dayKeys.length; j++) {
    var key = dayKeys[j];
    var dayTimestamps = dayGroups[key];
    dayTimestamps.sort(function (a, b) { return a - b; });
    var dayStats = calculateDayWorkTime(dayTimestamps);
    totalWorkMinutes += dayStats.workMinutes;
    totalBreakMinutes += dayStats.breakMinutes;
  }

  var dayCount = dayKeys.length;
  var avgDayMinutes = dayCount > 0 ? Math.round(totalWorkMinutes / dayCount) : 0;

  return {
    totalWork: formatMinutes(totalWorkMinutes),
    totalBreak: formatMinutes(totalBreakMinutes),
    avgDayDuration: formatMinutes(avgDayMinutes)
  };
}

function calculateDayWorkTime(timestamps) {
  if (!timestamps || !timestamps.length) {
    return {
      workTime: "0ч 0м",
      breakTime: "0ч 0м",
      shift: "-",
      forecast: "-",
      workMinutes: 0,
      breakMinutes: 0
    };
  }

  var now = new Date();
  var MAX_BREAK = 90;
  var ERROR_MARGIN = 15;
  var WORK_INTERVAL = 5;

  var firstTime = timestamps[0];
  var lastTime = timestamps[timestamps.length - 1];

  var workMinutes = 0;
  var breakMinutes = 0;

  var firstHour = firstTime.getHours();
  var shift = "09:00-18:00";
  var shiftEndHour = 18;
  if (firstHour >= 10 && firstHour < 14 || (firstHour >= 18 && firstHour < 21)) {
    shift = "11:00-20:00";
    shiftEndHour = 20;
  }

  for (var i = 0; i < timestamps.length - 1; i++) {
    var diff = (timestamps[i + 1] - timestamps[i]) / 60000;
    if (diff <= WORK_INTERVAL) {
      workMinutes += diff;
    } else if (diff <= (MAX_BREAK + ERROR_MARGIN)) {
      breakMinutes += diff;
    }
  }

  var isToday = (
    lastTime.getDate() === now.getDate() &&
    lastTime.getMonth() === now.getMonth() &&
    lastTime.getFullYear() === now.getFullYear()
  );

  var forecast = "-";
  if (isToday) {
    var minutesSinceLastFix = (now - lastTime) / 60000;
    if (minutesSinceLastFix <= (MAX_BREAK + ERROR_MARGIN)) {
      if (minutesSinceLastFix <= WORK_INTERVAL) {
        workMinutes += minutesSinceLastFix;
      } else {
        breakMinutes += minutesSinceLastFix;
      }

      var shiftEnd = new Date(firstTime);
      shiftEnd.setHours(shiftEndHour, 0, 0, 0);

      var totalElapsed = (now - firstTime) / 60000;
      var totalToShiftEnd = (shiftEnd - firstTime) / 60000;

      if (totalElapsed > 0 && workMinutes > 0) {
        var workRate = workMinutes / totalElapsed;
        var projectedWork = workRate * totalToShiftEnd;
        forecast = formatMinutes(Math.round(projectedWork));
      }
    } else {
      forecast = "День завершён";
    }
  }

  var roundedWork = Math.round(workMinutes);
  var roundedBreak = Math.round(breakMinutes);

  return {
    workTime: formatMinutes(roundedWork),
    breakTime: formatMinutes(roundedBreak),
    shift: shift,
    forecast: forecast,
    workMinutes: roundedWork,
    breakMinutes: roundedBreak
  };
}

function formatMinutes(minutes) {
  if (minutes < 0) minutes = 0;
  var hours = Math.floor(minutes / 60);
  var mins = minutes % 60;
  return hours + "ч " + mins + "м";
}

// =============================================================================
// 4. ЗАЩИТА ДАННЫХ (С ИСКЛЮЧЕНИЕМ КОЛОНОК H И I)
// =============================================================================

/**
 * Защищает строки с данными за прошлые дни от редактирования
 * ВАЖНО: Колонки H (ФИО) и I (Дата) остаются редактируемыми для всех пользователей!
 * Защищает все скрытые архивные листы
 */
function protectPastRows() {
  var startTime = new Date().getTime();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FIKSA");
  if (!sheet) {
    Logger.log("Лист FIKSA не найден");
    return;
  }

  // Удаляем старую защиту диапазона
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === PROTECTION_DESCRIPTION) {
      protections[i].remove();
    }
  }

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 2, lastRow - 1, 8);
    var allData = dataRange.getValues();
    
    var today = new Date();
    var startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var lockEndRow = 1;

    for (var i = 0; i < allData.length; i++) {
      var bValue = allData[i][0];
      var timeValue = allData[i][7];
      
      if (!bValue) continue;
      if (!timeValue) continue;
      
      var recordDate = timeValue instanceof Date ? timeValue : parseDateTime(timeValue);
      if (recordDate && recordDate < startOfToday) {
        lockEndRow = i + 2;
      }
    }

    if (lockEndRow > 1) {
      // ЗАЩИЩАЕМ ТОЛЬКО КОЛОНКИ A-G (не H и I!)
      var protection = sheet
        .getRange(2, 1, lockEndRow - 1, 7) // Колонки A-G
        .protect()
        .setDescription(PROTECTION_DESCRIPTION);

      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      protection.setWarningOnly(false);

      var editors = protection.getEditors();
      if (editors && editors.length) protection.removeEditors(editors);

      var me = Session.getEffectiveUser().getEmail();
      if (me) protection.addEditor(me);
      
      Logger.log("Защищено строк: " + (lockEndRow - 1) + " (колонки A-G, H и I остаются редактируемыми)");
    }
  }

  protectArchiveSheets(ss);
  
  var endTime = new Date().getTime();
  Logger.log("protectPastRows выполнен за " + (endTime - startTime) + " мс");
  
  ss.toast(
    "Защита данных обновлена за " + (endTime - startTime) + " мс\n" +
    "Колонки H (ФИО) и I (Дата) остаются редактируемыми",
    "✅ Готово",
    5
  );
}

/**
 * Защищает все скрытые архивные листы
 */
function protectArchiveSheets(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var me = Session.getEffectiveUser().getEmail();

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (!sheet.isSheetHidden()) continue;

    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var j = 0; j < protections.length; j++) {
      if (protections[j].getDescription() === PROTECTION_ARCHIVE_DESCRIPTION) {
        protections[j].remove();
      }
    }

    var protection = sheet.protect().setDescription(PROTECTION_ARCHIVE_DESCRIPTION);
    protection.setWarningOnly(false);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    var editors = protection.getEditors();
    if (editors && editors.length) protection.removeEditors(editors);

    if (me) protection.addEditor(me);
  }
}

// =============================================================================
// 5. АРХИВИРОВАНИЕ
// =============================================================================

/**
 * Переносит данные FIKSA в архивный лист
 * Создаёт скрытый лист с именем "ИмяОператора ММ.ГГГГ"
 * Очищает FIKSA и восстанавливает формулы
 * АВТОМАТИЧЕСКИ ЗАПУСКАЕТСЯ 19 ЧИСЛА КАЖДОГО МЕСЯЦА В 23:00
 */
function transferFiksa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  var sheetFiksa = ss.getSheetByName('FIKSA');
  if (!sheetFiksa) {
    ss.toast('Лист "FIKSA" не найден.', '❌ Ошибка', 5);
    Logger.log('transferFiksa: Лист FIKSA не найден');
    return;
  }
  var sheetSetting = ss.getSheetByName('SETTING');
  if (!sheetSetting) {
    ss.toast('Лист "SETTING" не найден.', '❌ Ошибка', 5);
    Logger.log('transferFiksa: Лист SETTING не найден');
    return;
  }

  var baseName = (sheetSetting.getRange('B2').getValue() || '').toString().trim();
  if (!baseName) {
    ss.toast('SETTING!B2 пуста. Укажите имя-основу.', '❌ Ошибка', 5);
    Logger.log('transferFiksa: SETTING!B2 пуст');
    return;
  }

  var monthYear = Utilities.formatDate(new Date(), tz, 'MM.yyyy');
  var newSheetName = sanitizeSheetName(baseName + ' ' + monthYear);

  var copied;
  try {
    copied = sheetFiksa.copyTo(ss);
  } catch (e) {
    ss.toast('Ошибка при копировании: ' + e.message, '❌ Ошибка', 5);
    Logger.log('transferFiksa: Ошибка копирования - ' + e.message);
    throw e;
  }

  var existing = ss.getSheetByName(newSheetName);
  if (existing) {
    ss.deleteSheet(copied);
    ss.toast(
      'Архив "' + newSheetName + '" уже существует. Переименуйте старый архив.',
      '❌ Ошибка',
      5
    );
    Logger.log('transferFiksa: Архив ' + newSheetName + ' уже существует');
    return;
  }

  try {
    copied.setName(newSheetName);
  } catch (e) {
    ss.deleteSheet(copied);
    ss.toast('Не удалось переименовать копию в "' + newSheetName + '": ' + e.message, '❌ Ошибка', 5);
    Logger.log('transferFiksa: Ошибка переименования - ' + e.message);
    return;
  }

  try {
    ss.setActiveSheet(copied);
    ss.moveActiveSheet(ss.getNumSheets());
  } catch (err) {
    Logger.log('transferFiksa move sheet: ' + err);
  }

  freezeLookupColumns(copied);

  ss.setActiveSheet(sheetFiksa);
  copied.hideSheet();

  var lastRow = sheetFiksa.getLastRow();
  if (lastRow >= 2) {
    var startRow = 2;
    var startCol = 2;
    var endCol = 10;
    var numRows = lastRow - startRow + 1;
    var numCols = endCol - startCol + 1;
    sheetFiksa.getRange(startRow, startCol, numRows, numCols).clearContent();
  }

  applyLookupFormulas(sheetFiksa);

  ss.toast(
    'Данные перенесены на "' + newSheetName + '", архив скрыт, FIKSA очищен.',
    '✅ Архивация завершена',
    5
  );
  
  Logger.log('transferFiksa: Архивация успешно завершена. Создан лист: ' + newSheetName);
}

/**
 * Замораживает LOOKUP формулы в архиве (превращает в значения)
 */
function freezeLookupColumns(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var height = lastRow - 1;
  sheet.getRange(2, 3, height, 2).copyTo(
    sheet.getRange(2, 3, height, 2),
    SpreadsheetApp.CopyPasteType.PASTE_VALUES,
    false
  );
}

/**
 * Применяет LOOKUP формулы в колонках C и D
 */
function applyLookupFormulas(sheet) {
  if (sheet.getMaxRows() < 2) sheet.insertRowsAfter(1, 1);

  var maxRows = sheet.getMaxRows();
  var fillRows = Math.max(maxRows - 1, 1);

  sheet.getRange('C2').setFormula(LOOKUP_FORMULA_C);
  sheet.getRange('D2').setFormula(LOOKUP_FORMULA_D);

  if (fillRows > 1) {
    sheet.getRange('C2').copyTo(
      sheet.getRange(3, 3, fillRows - 1, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
      false
    );
    sheet.getRange('D2').copyTo(
      sheet.getRange(3, 4, fillRows - 1, 1),
      SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
      false
    );
  }
}

// =============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Управляет видимостью листов статистики
 * Оставляет видимыми: FIKSA, Аризалар, SETTING, текущий месяц, предыдущий месяц
 * Скрывает все остальные листы со статистикой
 */
function manageSheetVisibility(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  
  // Листы, которые ВСЕГДА должны быть видимыми
  var alwaysVisible = [
    "FIKSA",
    "Аризалар",
    "Статистика",      // Текущий месяц (для скрипта сбора)
    "Предыдущий месяц", // Предыдущий месяц (для скрипта сбора)
    "Сводка по дням"   // Сводная статистика по всем операторам
  ];
  
  var sheets = ss.getSheets();
  var hiddenCount = 0;
  var shownCount = 0;
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Проверяем, должен ли лист быть видимым
    var shouldBeVisible = false;
    for (var j = 0; j < alwaysVisible.length; j++) {
      if (sheetName === alwaysVisible[j]) {
        shouldBeVisible = true;
        break;
      }
    }
    
    if (shouldBeVisible) {
      // Показываем лист, если он был скрыт
      if (sheet.isSheetHidden()) {
        sheet.showSheet();
        shownCount++;
        Logger.log("Показан лист: " + sheetName);
      }
    } else {
      // Скрываем архивные листы (месяц + год в названии, но не текущие)
      var isArchive = /\d{2}\.\d{4}/.test(sheetName) || sheetName.indexOf("📊") === 0;
      if (isArchive && !sheet.isSheetHidden()) {
        sheet.hideSheet();
        hiddenCount++;
        Logger.log("Скрыт лист: " + sheetName);
      }
    }
  }
  
  if (hiddenCount > 0 || shownCount > 0) {
    Logger.log("manageSheetVisibility: скрыто " + hiddenCount + ", показано " + shownCount);
  }
}

/**
 * Очищает имя листа от запрещённых символов
 */
function sanitizeSheetName(name) {
  if (!name) return 'Sheet';
  var cleaned = name.replace(/[\/\\\?\*\[\]\:]/g, '').slice(0, 100).trim();
  return cleaned || 'Sheet';
}

// =============================================================================
// ФУНКЦИИ ДЛЯ ВНЕШНЕГО ДОСТУПА К ДАННЫМ (для скрипта-сборщика)
// =============================================================================

/**
 * Возвращает сводные данные для внешнего скрипта-сборщика
 * Этот метод можно вызвать из другой таблицы для сбора статистики
 */
function getStatisticsSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName("SETTING");
  
  if (!settingSheet) {
    return {
      error: "SETTING sheet not found",
      operatorName: "Unknown"
    };
  }
  
  var operatorName = (settingSheet.getRange("B2").getValue() || "").toString().trim();
  
  return {
    operatorName: operatorName,
    currentMonth: getSheetSummary("Статистика"),
    previousMonth: getSheetSummary("Предыдущий месяц"),
    lastUpdated: new Date()
  };
}

/**
 * Получает сводку с листа статистики
 */
function getSheetSummary(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return {
      error: "Sheet not found: " + sheetName,
      exists: false
    };
  }
  
  try {
    // Читаем ключевые ячейки из листа статистики
    var period = sheet.getRange("A2").getValue(); // Период: дата - дата
    var updated = sheet.getRange("A3").getValue(); // Обновлено: дата время
    
    var totalFixes = sheet.getRange("B6").getValue() || 0; // Всего фиксаций
    var totalUnique = sheet.getRange("B7").getValue() || 0; // Уникальных записей
    var totalOpen = sheet.getRange("B8").getValue() || 0; // Открыто
    var totalClosed = sheet.getRange("B9").getValue() || 0; // Закрыто
    
    // Статусы (строки 13-17)
    var status1 = sheet.getRange("B13").getValue() || 0; // отрицательный
    var status2 = sheet.getRange("B14").getValue() || 0; // положительный
    var status3 = sheet.getRange("B15").getValue() || 0; // заявка закрыта
    var status4 = sheet.getRange("B16").getValue() || 0; // открыть карту
    var status5 = sheet.getRange("B17").getValue() || 0; // тиббиёт ходими
    
    return {
      exists: true,
      period: String(period),
      updated: String(updated),
      totalFixes: totalFixes,
      totalUnique: totalUnique,
      totalOpen: totalOpen,
      totalClosed: totalClosed,
      statuses: {
        negative: status1,
        positive: status2,
        callFailed: status3,
        openCard: status4,
        medical: status5
      }
    };
  } catch (err) {
    Logger.log("Ошибка чтения сводки из " + sheetName + ": " + err);
    return {
      error: String(err),
      exists: true
    };
  }
}

/**
 * Принудительно обновляет статистику (для вызова извне)
 */
function forceUpdateStatistics() {
  Logger.log("forceUpdateStatistics вызван извне");
  updateCurrentMonthStatistics();
  return {
    success: true,
    message: "Statistics updated",
    timestamp: new Date()
  };
}

// =============================================================================
// ЗАПОЛНЕНИЕ АРХИВНЫХ ДАННЫХ (ДО 20 СЕНТЯБРЯ 2024)
// =============================================================================

/**
 * Автоматическое заполнение архивных данных (запускается по триггеру)
 */
function fillArchiveDataAuto() {
  var startTime = new Date().getTime();
  
  try {
    var result = fillArchiveData();
    
    var endTime = new Date().getTime();
    Logger.log("fillArchiveDataAuto завершено: листов " + result.sheetsProcessed + 
               ", строк " + result.rowsFilled + ", время " + 
               Math.round((endTime - startTime) / 1000) + " сек");
  } catch (err) {
    Logger.log("Ошибка в fillArchiveDataAuto: " + err);
  }
}

/**
 * Заполняет ФИО и дату/время в архивных листах (до 20.09.2024)
 * Логика для старых данных: заполняет каждую строку с номером карты
 */
function fillArchiveData(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = Session.getScriptTimeZone();
  
  // Граница периода: 20 сентября 2024, 00:00:00
  var cutoffDate = new Date(2024, 8, 20, 0, 0, 0); // месяц 8 = сентябрь (0-indexed)
  
  var sheetsProcessed = 0;
  var totalRowsFilled = 0;
  
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Обрабатываем только архивные листы с датой в названии
    if (!/\d{2}\.\d{4}/.test(sheetName)) continue;
    
    Logger.log("Обработка листа: " + sheetName);
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("  Пропуск: нет данных");
      continue;
    }
    
    // Читаем все данные
    var dataRange = sheet.getRange(2, 2, lastRow - 1, 8); // B-I (столбцы 2-9)
    var allData = dataRange.getValues();
    
    var previousFio = "";
    var previousDate = null;
    var changedH = false;
    var changedI = false;
    var rowsFilledInSheet = 0;
    
    for (var r = 0; r < allData.length; r++) {
      var colB = allData[r][0];  // B - номер карты
      var colH = allData[r][6];  // H - ФИО
      var colI = allData[r][7];  // I - дата/время
      
      // Если строка пустая (нет номера карты), пропускаем
      if (!colB || String(colB).trim() === "") continue;
      
      // Проверяем дату записи - если после 20.09.2024, пропускаем эту строку
      var recordDate = parseDateTime(colI);
      if (recordDate && recordDate >= cutoffDate) {
        // Обновляем previousDate для следующих строк
        if (recordDate) previousDate = recordDate;
        continue;
      }
      
      // ОБРАБОТКА КОЛОНКИ I (дата/время)
      if (colI && colI !== "") {
        // Дата есть - сохраняем как предыдущую
        previousDate = colI;
      } else if (previousDate) {
        // Дата пустая - заполняем предыдущей
        allData[r][7] = previousDate;
        changedI = true;
        rowsFilledInSheet++;
      }
      
      // ОБРАБОТКА КОЛОНКИ H (ФИО)
      var hVal = String(colH).trim();
      var isEmpty = !colH || hVal === "";
      var isError = hVal.indexOf("#") === 0 || 
                    hVal.toLowerCase().indexOf("error") !== -1 || 
                    hVal.toLowerCase().indexOf("ошибка") !== -1 ||
                    hVal.toLowerCase().indexOf("ref") !== -1 ||
                    hVal === "REF!" ||
                    hVal === "#REF!";
      
      if (!isEmpty && !isError) {
        // ФИО корректное - сохраняем как предыдущее
        previousFio = colH;
      } else if (previousFio !== "" && (isEmpty || isError)) {
        // ФИО пустое или ошибка - заполняем предыдущим
        allData[r][6] = previousFio;
        changedH = true;
        if (!changedI) rowsFilledInSheet++; // считаем только если еще не посчитали
      }
    }
    
    // Записываем изменения, если они были
    if (changedH || changedI) {
      var updateData = allData.map(function(row) { return [row[6], row[7]]; });
      sheet.getRange(2, 8, updateData.length, 2).setValues(updateData);
      
      sheetsProcessed++;
      totalRowsFilled += rowsFilledInSheet;
      
      Logger.log("  ✓ Лист обработан: " + rowsFilledInSheet + " строк заполнено");
    } else {
      Logger.log("  Изменений не требуется");
    }
  }
  
  Logger.log("fillArchiveData завершено: листов " + sheetsProcessed + ", строк " + totalRowsFilled);
  
  return {
    success: true,
    message: "Архивные данные заполнены",
    sheetsProcessed: sheetsProcessed,
    rowsFilled: totalRowsFilled
  };
}

// =============================================================================
// КОНЕЦ ПОЛНОГО СКРИПТА
// =============================================================================

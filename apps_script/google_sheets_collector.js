/**
 * Универсальная отправка сообщения в Telegram
 */
function sendTelegramMessage(message) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  var payload = {
    chat_id: TELEGRAM_CHAT_ID,
    text: message,
    parse_mode: "Markdown"
  };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Telegram отправлено: " + response.getContentText());
  } catch (e) {
    Logger.log("Ошибка отправки в Telegram: " + e);
  }
}

/**
 * Отправляет красивый отчет из листа "Сводка по дням текущего месяца" в Telegram
 */
function sendCurrentMonthSummarySheetToTelegram() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Сводка по дням текущего месяца");
  if (!sheet) {
    Logger.log("Лист 'Сводка по дням текущего месяца' не найден");
    return;
  }
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 6) {
    Logger.log("Недостаточно данных для отчета");
    return;
  }

  var header = rows[0][0] || "";
  var period = rows[1][0] || "";
  var source = rows[2][0] || "";
  var tableHeader = rows[3];
  var message = "*" + header + "*\n" + period + "\n" + source + "\n\n";
  message += "_" + tableHeader[0] + "_ | _" + tableHeader[1] + "_\n";

  var total = "";
  for (var i = 4; i < rows.length; i++) {
    var fio = rows[i][0];
    var closed = rows[i][1];
    if (fio && fio !== "ИТОГО:") {
      message += "*" + fio + "*: " + closed + "\n";
    }
    if (fio === "ИТОГО:") {
      total = closed;
    }
  }
  if (total !== "") {
    message += "\n*ИТОГО*: " + total + "\n";
  }
  sendTelegramMessage(message);
}
/**
 * Создает лист "Сводка по дням текущего месяца" в формате отчета
 */
function createCurrentMonthDailySummarySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(DAILY_SUMMARY_SHEET_NAME);
  if (!sourceSheet) {
    Logger.log("Лист 'Сводка по дням' не найден");
    return;
  }

  var today = new Date();
  var currentMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM.yyyy");
  var currentDay = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd.MM.yyyy");
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Нет данных для формирования отчета");
    return;
  }
  var data = sourceSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var monthRows = [];
  var closedByOperator = {};
  var operators = {};
  for (var i = 0; i < data.length; i++) {
    var dateStr = String(data[i][0] || "").trim();
    var match = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (match) {
      var rowMonth = match[2] + "." + match[3];
      if (rowMonth === currentMonth) {
        var fio = String(data[i][1] || "").trim();
        var closed = parseInt(data[i][4]) || 0;
        if (!operators[fio]) operators[fio] = 0;
        operators[fio] += closed;
      }
    }
  }
  var operatorList = Object.keys(operators);
  if (operatorList.length === 0) {
    Logger.log("Нет статистики по дням за текущий месяц");
    return;
  }

  // Создаем/обновляем лист
  var sheetName = "Сводка по дням текущего месяца";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  // Формируем отчет
  var periodStart = null;
  var periodEnd = null;
  for (var i = 0; i < data.length; i++) {
    var dateStr = String(data[i][0] || "").trim();
    var match = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
    if (match && (match[2] + "." + match[3]) === currentMonth) {
      var dateObj = new Date(match[3], match[2] - 1, match[1]);
      if (!periodStart || dateObj < periodStart) periodStart = dateObj;
      if (!periodEnd || dateObj > periodEnd) periodEnd = dateObj;
    }
  }
  var periodStr = periodStart && periodEnd
    ? Utilities.formatDate(periodStart, Session.getScriptTimeZone(), "dd.MM.yyyy") + " - " + Utilities.formatDate(periodEnd, Session.getScriptTimeZone(), "dd.MM.yyyy")
    : "";

  var rows = [];
  rows.push(["ОТЧЕТ ЗА ДЕНЬ: " + currentDay + " (ТОЛЬКО ЗАКРЫТЫЕ)", ""]);
  rows.push(["Текущий цикл: Период: " + periodStr, ""]);
  rows.push(["Данные из таблицы 'СТАТИСТИКА ПО ДНЯМ'", ""]);
  rows.push(["Филиал", "Закрыто (" + currentDay + ")"]);

  var total = 0;
  for (var i = 0; i < operatorList.length; i++) {
    var fio = operatorList[i];
    var closed = operators[fio];
    rows.push([fio, closed]);
    total += closed;
  }
  // Добавляем пустые строки для визуального разделения, если нужно
  for (var i = rows.length; i < 40; i++) {
    rows.push(["", ""]);
  }
  rows.push(["ИТОГО:", total]);

  // Записываем в лист
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // Форматирование
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setFontSize(12);
  sheet.getRange(2, 1, 1, 2).setFontStyle("italic");
  sheet.getRange(3, 1, 1, 2).setBackground("#d9ead3");
  sheet.getRange(4, 1, 1, 2).setFontWeight("bold").setBackground("#b6d7a8");
  sheet.getRange(rows.length, 1, 1, 2).setFontWeight("bold").setBackground("#f3f3f3");
  sheet.autoResizeColumns(1, 2);
  sheet.setFrozenRows(4);

  Logger.log("Лист 'Сводка по дням текущего месяца' обновлен");
}
/**
 * Отправляет статистику по дням за текущий месяц в Telegram
 */
function sendCurrentMonthDailySummaryToTelegram() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DAILY_SUMMARY_SHEET_NAME);
  if (!sheet) {
    Logger.log("Лист 'Сводка по дням' не найден");
    return;
  }
  var today = new Date();
  var currentMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM.yyyy");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Нет данных для отправки в Telegram");
    return;
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var monthRows = data.filter(function(row) {
    var dateStr = String(row[0] || "").trim();
    var match = dateStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
    return match && (match[2] + "." + match[3]) === currentMonth;
  });
  if (monthRows.length === 0) {
    Logger.log("Нет статистики по дням за текущий месяц");
    return;
  }
  var message = "Статистика по дням за " + currentMonth + ":\n";
  message += "Дата | ФИО | Фиксаций | Уникальных | Закрытых | Открытых | Повторных\n";
  monthRows.forEach(function(row) {
    message += row.join(" | ") + "\n";
  });
  sendTelegramMessage(message);
}
/**
 * =============================================================================
 * СКРИПТ СБОРА СТАТИСТИКИ СО ВСЕХ ТАБЛИЦ ОПЕРАТОРОВ
 * =============================================================================
 * Версия: 3.0 ОПТИМИЗИРОВАННАЯ
 * Дата: 01.12.2025
 * 
 * ✨ ОСНОВНЫЕ УЛУЧШЕНИЯ В ВЕРСИИ 3.0:
 * - Упрощённая архитектура без сложных очередей
 * - Улучшенная функция чтения данных с динамическим поиском
 * - Детальное логирование каждого шага
 * - Автоматическая диагностика при проблемах с данными
 * - Код стал в 3 раза проще и понятнее
 * 
 * 📋 ФУНКЦИОНАЛ:
 * - Собирает данные с листов "Статистика" всех операторов
 * - Создаёт сводные листы: Текущий месяц, Предыдущий месяц, Сводка по дням
 * - Автоматическая отправка в Telegram (13:00 и 18:00)
 * - Работает по триггерам каждые 2 часа
 * 
 * 🚀 БЫСТРЫЙ СТАРТ:
 * 1. Создайте новую Google Таблицу для сбора статистики
 * 2. Откройте Расширения → Apps Script
 * 3. Вставьте этот скрипт и сохраните
 * 4. Обновите страницу таблицы - появится меню "📊 Сбор статистики"
 * 5. Заполните лист "Настройки" (ID таблиц операторов + статус "активен")
 * 6. Запустите ОДИН РАЗ: Меню → ⚙️ Настроить автосбор
 * 7. Протестируйте: Меню → 🔄 Обновить все данные
 * 
 * 📝 ВАЖНО:
 * - У вас должен быть доступ ко всем таблицам операторов (Редактор или Читатель)
 * - В каждой таблице оператора должен быть лист "Статистика"
 * - Смотрите логи: Apps Script → Журнал выполнения
 * 
 * 🔧 ОТЛАДКА:
 * - Читайте файл ОТЛАДКА.md для решения проблем
 * - Все подробности в файле ИСПРАВЛЕНИЯ.md
 * =============================================================================
 */

// =============================================================================

/*******************  НАСТРОЙКИ TELEGRAM  *******************/
const TELEGRAM_TOKEN = '7940976522:AAEFqt3QwaoOtPqqwZmCcJoDw9e3RVuPyq8';
const TELEGRAM_CHAT_ID = '2012682567';
// ГЛОБАЛЬНЫЕ НАСТРОЙКИ
// =============================================================================

/**
 * Структура листа "Настройки":
 * 
 * A          | B                                      | C
 * ФИО        | ID таблицы                             | Статус
 * ---------- | -------------------------------------- | --------
 * Оператор 1 | 1abc...xyz                             | активен
 * Оператор 2 | 2def...uvw                             | активен
 */

var SETTINGS_SHEET_NAME = "Настройки";
var CURRENT_STATS_SHEET_NAME = "Текущий месяц - Сводка";
var PREVIOUS_STATS_SHEET_NAME = "Предыдущий месяц - Сводка";
var DAILY_SUMMARY_SHEET_NAME = "Сводка по дням";
var MONTHLY_STATS_PREFIX = "📊 "; // Префикс для помесячных листов

// Настройки оптимизации
var MAX_EXECUTION_TIME = 300000; // 5 минут в миллисекундах
var BATCH_SIZE = 3; // Обрабатываем по 3 оператора за раз
var CACHE_DURATION = 21600; // Кэш на 6 часов
var QUEUE_SHEET_NAME = "_Очередь_"; // Служебный лист для очереди задач

// =============================================================================
// НАСТРОЙКА ТРИГГЕРОВ
// =============================================================================

/**
 * Настраивает автоматические триггеры сбора данных
 * ЗАПУСТИТЬ ОДИН РАЗ!
 */
function setupCollectorTriggers() {
  Logger.log("Настройка триггеров...");
  
  // Удаляем ВСЕ старые триггеры
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log("Найдено старых триггеров: " + triggers.length);
  
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log("Старые триггеры удалены");
  
  // Триггер сбора каждые 2 часа
  ScriptApp.newTrigger("collectAllStatistics")
    .timeBased()
    .everyHours(2)
    .create();

  // Триггер отправки статистики по дням в Telegram в 13:00
  ScriptApp.newTrigger("sendDailySummaryToTelegram")
    .timeBased()
    .atHour(13)
    .everyDays(1)
    .create();

  // Триггер отправки статистики по дням в Telegram в 18:00
  ScriptApp.newTrigger("sendDailySummaryToTelegram")
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();

  Logger.log("✓ Триггеры Telegram (13:00, 18:00) созданы");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Автоматический сбор настроен!\n\n" +
    "Статистика будет обновляться каждые 2 часа\n" +
    "Telegram отчет по дням: 13:00 и 18:00",
    "Настройка завершена",
    10
  );
}

// =============================================================================
// МЕНЮ
// =============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("📊 Сбор статистики")
    .addItem("🔄 Обновить все данные", "collectAllStatistics")
    .addItem("📈 Собрать только текущий месяц", "collectCurrentMonth")
    .addItem("📉 Собрать только предыдущий месяц", "collectPreviousMonth")
    .addSeparator()
    .addItem("📅 Создать помесячную статистику", "createMonthlyStatistics")
    .addSeparator()
    .addItem("📤 Отправить по дням за текущий месяц в Telegram", "sendCurrentMonthDailySummaryToTelegram")
    .addItem("📋 Сформировать лист сводки по дням текущего месяца", "createCurrentMonthDailySummarySheet")
    .addItem("📤 Отправить отчет по дням текущего месяца в Telegram", "sendCurrentMonthSummarySheetToTelegram")
    .addSeparator()
    .addItem("⚙️ Настроить автосбор", "setupCollectorTriggers")
    .addToUi();
}

// =============================================================================
// ОСНОВНЫЕ ФУНКЦИИ СБОРА
// =============================================================================

/**
 * Отправляет статистику по дням в Telegram
 */
function sendDailySummaryToTelegram() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DAILY_SUMMARY_SHEET_NAME);
  if (!sheet) {
    Logger.log("Лист 'Сводка по дням' не найден");
    return;
  }
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Нет статистики по дням за сегодня. Запускаем сбор...");
    try {
      collectDailySummary();
    } catch (e) {
      Logger.log("Ошибка при сборе сводки по дням: " + e);
    }
    lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Нет статистики по дням за сегодня даже после обновления.");
      return;
    }
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var summaryRows = data.filter(function(row) {
    return String(row[0]) === today;
  });
  if (summaryRows.length === 0) {
    Logger.log("Нет статистики по дням за сегодня.");
    return;
  }
  var message = "Статистика по дням за " + today + ":\n";
  message += "Дата | ФИО | Фиксаций | Уникальных | Закрытых | Открытых | Повторных\n";
  summaryRows.forEach(function(row) {
    message += row.join(" | ") + "\n";
  });
  sendTelegramMessage(message);
}

/**
 * УПРОЩЕННАЯ версия сбора всех данных - без сложных очередей
 */
function collectAllStatistics() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("\n========================================");
  Logger.log("НАЧАЛО СБОРА СТАТИСТИКИ");
  Logger.log("Время: " + new Date().toLocaleString());
  Logger.log("========================================\n");

  try {
    // 1. Собираем текущий месяц
    Logger.log("[1/3] Сбор текущего месяца...");
    collectCurrentMonth();
    
    // 2. Собираем предыдущий месяц
    Logger.log("[2/3] Сбор предыдущего месяца...");
    collectPreviousMonth();
    
    // 3. Собираем сводку по дням
    Logger.log("[3/3] Сбор сводки по дням...");
    collectDailySummary();

    var endTime = new Date().getTime();
    var duration = Math.round((endTime - startTime) / 1000);

    Logger.log("\n========================================");
    Logger.log("✅ СБОР ЗАВЕРШЕН ЗА " + duration + " СЕК");
    Logger.log("========================================\n");

    ss.toast(
      "✅ Все данные собраны!\n" +
      "Время: " + duration + " сек",
      "Готово",
      5
    );

  } catch (err) {
    Logger.log("\n❌ КРИТИЧЕСКАЯ ОШИБКА: " + err);
    Logger.log("Stack: " + err.stack);
    ss.toast("❌ Ошибка: " + err, "Ошибка", 10);
  }
}

// УПРОЩЕННАЯ ВЕРСИЯ - без сложных очередей и таймаутов
// Все функции работают синхронно и последовательно

/**
 * ОПТИМИЗИРОВАННЫЙ сбор данных текущего месяца
 * Быстрая обработка с минимальным логированием
 */
function collectCurrentMonth() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var operators = getOperatorList();

  if (operators.length === 0) {
    Logger.log("⚠ Нет операторов в 'Настройки'");
    ss.toast("Нет операторов в листе 'Настройки'", "Ошибка", 3);
    return;
  }

  Logger.log("\n▶ Сбор ТЕКУЩЕГО месяца: " + operators.length + " операторов");

  // Подготовка листа
  var sheet = ss.getSheetByName(CURRENT_STATS_SHEET_NAME) || ss.insertSheet(CURRENT_STATS_SHEET_NAME);
  sheet.clear();

  var headers = [
    "ФИО оператора", "Дата обновления", "Период",
    "Всего фиксаций", "Уникальных записей", "Открыто", "Закрыто",
    "Отрицательный", "Положительный", "Заявка закрыта", "Открыть карту", "Тиббиёт ходими"
  ];

  var allData = [headers];
  var success = 0, fail = 0;
  
  // Быстрая обработка
  for (var i = 0; i < operators.length; i++) {
    var op = operators[i];
    
    if (op.status.toLowerCase() !== "активен") continue;

    try {
      var data = collectStatsFromSheet(op.spreadsheetId, "Статистика", op.name);
      if (data) {
        allData.push(data);
        success++;
      } else {
        fail++;
      }
    } catch (err) {
      Logger.log("  ✗ " + op.name + ": " + err.message);
      fail++;
    }
  }

  // Записываем данные одним блоком (быстро)
  if (allData.length > 1) {
    sheet.getRange(1, 1, allData.length, headers.length).setValues(allData);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  var duration = Math.round((new Date().getTime() - startTime) / 1000);
  Logger.log("✓ Текущий месяц: " + success + " успешно, " + fail + " ошибок (" + duration + " сек)\n");
}

/**
 * ОПТИМИЗИРОВАННЫЙ сбор предыдущего месяца
 */
function collectPreviousMonth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var operators = getOperatorList();
  
  Logger.log("▶ Сбор ПРЕДЫДУЩЕГО месяца...");
  
  var sheet = ss.getSheetByName(PREVIOUS_STATS_SHEET_NAME) || ss.insertSheet(PREVIOUS_STATS_SHEET_NAME);
  sheet.clear();
  
  var headers = [
    "ФИО оператора", "Дата обновления", "Период",
    "Всего фиксаций", "Уникальных записей", "Открыто", "Закрыто",
    "Отрицательный", "Положительный", "Заявка закрыта", "Открыть карту", "Тиббиёт ходими"
  ];
  
  var allData = [headers];
  var count = 0;
  
  for (var i = 0; i < operators.length; i++) {
    if (operators[i].status.toLowerCase() !== "активен") continue;
    
    try {
      var data = collectStatsFromSheet(operators[i].spreadsheetId, "Предыдущий месяц", operators[i].name);
      if (data) {
        allData.push(data);
        count++;
      }
    } catch (err) {
      Logger.log("  ✗ " + operators[i].name + ": " + err.message);
    }
  }
  
  if (allData.length > 1) {
    sheet.getRange(1, 1, allData.length, headers.length).setValues(allData);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  
  Logger.log("✓ Предыдущий месяц: " + count + " записей\n");
}

/**
 * УЛУЧШЕННАЯ сводка по дням - табличный формат с группировкой по датам
 * Формат: Дата | ФИО | Фиксаций | Уникальных | Закрытых | Открытых | Повторных
 */
function collectDailySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var operators = getOperatorList();
  
  Logger.log("▶ Сбор СВОДКИ по дням (табличный формат)...");
  
  var sheet = ss.getSheetByName(DAILY_SUMMARY_SHEET_NAME) || ss.insertSheet(DAILY_SUMMARY_SHEET_NAME);
  sheet.clear();
  
  // Собираем все данные с группировкой по датам
  var dataByDate = {};
  var allDates = [];
  
  for (var i = 0; i < operators.length; i++) {
    if (operators[i].status.toLowerCase() !== "активен") continue;
    
    try {
      var remote = SpreadsheetApp.openById(operators[i].spreadsheetId).getSheetByName("Сводка по дням");
      if (!remote || remote.getLastRow() < 2) continue;
      
      var data = remote.getRange(2, 1, remote.getLastRow() - 1, 7).getValues();
      for (var j = 0; j < data.length; j++) {
        if (data[j][0]) {
          var dateStr = String(data[j][0]).trim();
          
          // Группируем по датам
          if (!dataByDate[dateStr]) {
            dataByDate[dateStr] = [];
            allDates.push(dateStr);
          }
          
          dataByDate[dateStr].push({
            fio: data[j][1],
            fixes: data[j][2] || 0,
            unique: data[j][3] || 0,
            closed: data[j][4] || 0,
            open: data[j][5] || 0,
            repeated: data[j][6] || 0
          });
        }
      }
    } catch (err) {
      Logger.log("  ✗ " + operators[i].name + ": " + err.message);
    }
  }
  
  // Сортируем даты (новые сверху)
  allDates.sort(function(a, b) {
    return parseDate(b) - parseDate(a);
  });
  
  // Формируем итоговую таблицу
  var tableData = [];
  
  // Заголовок
  tableData.push([
    "Дата", 
    "ФИО оператора", 
    "Всего фиксаций", 
    "Уникальных карт", 
    "Закрыто", 
    "Открыто", 
    "Повторных"
  ]);
  
  // Данные по каждой дате с суммами
  var totalRows = 0;
  for (var d = 0; d < allDates.length; d++) {
    var dateStr = allDates[d];
    var operators = dataByDate[dateStr];
    
    // Сортируем операторов по убыванию закрытых
    operators.sort(function(a, b) {
      return (b.closed || 0) - (a.closed || 0);
    });
    
    // Суммы по дате
    var dayTotals = {
      fixes: 0,
      unique: 0,
      closed: 0,
      open: 0,
      repeated: 0
    };
    
    // Добавляем строки для каждого оператора
    for (var o = 0; o < operators.length; o++) {
      var op = operators[o];
      tableData.push([
        dateStr,
        op.fio,
        op.fixes,
        op.unique,
        op.closed,
        op.open,
        op.repeated
      ]);
      totalRows++;
      
      // Накапливаем суммы
      dayTotals.fixes += op.fixes || 0;
      dayTotals.unique += op.unique || 0;
      dayTotals.closed += op.closed || 0;
      dayTotals.open += op.open || 0;
      dayTotals.repeated += op.repeated || 0;
    }
    
    // Добавляем итоговую строку по дате
    tableData.push([
      dateStr,
      "ИТОГО за день",
      dayTotals.fixes,
      dayTotals.unique,
      dayTotals.closed,
      dayTotals.open,
      dayTotals.repeated
    ]);
    
    // Добавляем пустую строку между датами
    if (d < allDates.length - 1) {
      tableData.push(["", "", "", "", "", "", ""]);
    }
  }
  
  // Записываем в лист
  if (tableData.length > 1) {
    sheet.getRange(1, 1, tableData.length, 7).setValues(tableData);
    
    // Форматирование заголовка
    sheet.getRange(1, 1, 1, 7)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    
    // Форматирование столбцов с числами (выравнивание по центру)
    if (tableData.length > 1) {
      sheet.getRange(2, 3, tableData.length - 1, 5).setHorizontalAlignment("center");
    }
    
    // Форматируем итоговые строки (содержат "ИТОГО за день")
    for (var r = 2; r <= tableData.length; r++) {
      if (String(tableData[r-1][1]).indexOf("ИТОГО") !== -1) {
        sheet.getRange(r, 1, 1, 7)
          .setFontWeight("bold")
          .setBackground("#fff2cc")
          .setFontStyle("italic");
      }
    }
    
    // Замораживаем заголовок
    sheet.setFrozenRows(1);
    
    // Автоматическая ширина столбцов
    sheet.autoResizeColumns(1, 7);
    
    // Устанавливаем границы для таблицы
    sheet.getRange(1, 1, tableData.length, 7).setBorder(
      true, true, true, true, true, true,
      "#000000", SpreadsheetApp.BorderStyle.SOLID
    );
    
    Logger.log("✓ Сводка по дням: " + totalRows + " записей, дат: " + allDates.length + "\n");
  } else {
    Logger.log("⚠ Нет данных для сводки по дням\n");
  }
}

// =============================================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =============================================================================

/**
 * Получает список операторов из листа "Настройки"
 */
function getOperatorList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!sheet) {
    // Создаем лист настроек с примером
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange("A1:C1").setValues([["ФИО оператора", "ID таблицы", "Статус"]]);
    sheet.getRange("A2:C3").setValues([
      ["Иванов Иван", "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ", "активен"],
      ["Петров Петр", "ВСТАВЬТЕ_ID_ТАБЛИЦЫ_ЗДЕСЬ", "неактивен"]
    ]);
    
    sheet.getRange("A1:C1")
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");
    
    sheet.autoResizeColumns(1, 3);
    
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
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
 * Обновляет статистику в удаленной таблице оператора
 * УЛУЧШЕНО: проверка доступа с повторными попытками
 */
function updateRemoteStatistics(spreadsheetId, operatorName) {
  try {
    Logger.log("Проверка доступа к таблице " + operatorName + "...");
    
    // Попытка открыть таблицу с повторами
    var maxRetries = 3;
    var remoteSpreadsheet = null;
    
    for (var attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        remoteSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
        if (remoteSpreadsheet) {
          Logger.log("Успешное подключение к таблице " + operatorName);
          break;
        }
      } catch (retryErr) {
        Logger.log("Попытка " + attempt + "/" + maxRetries + " неудачна для " + operatorName + ": " + retryErr);
        if (attempt < maxRetries) {
          Utilities.sleep(2000); // Ждем 2 секунды перед повтором
        }
      }
    }
    
    if (!remoteSpreadsheet) {
      Logger.log("Не удалось открыть таблицу " + operatorName + " после " + maxRetries + " попыток");
      return false;
    }
    
    Logger.log("Статистика для " + operatorName + " обновляется автоматически (hourly trigger)");
    return true;
    
  } catch (err) {
    Logger.log("ОШИБКА при обновлении статистики для " + operatorName + ": " + err);
    return false;
  }
}

/**
 * МАКСИМАЛЬНО ОПТИМИЗИРОВАННАЯ версия сбора данных
 * Поддерживает разные форматы, быстро работает, надёжно читает
 */
function collectStatsFromSheet(spreadsheetId, sheetName, operatorName) {
  try {
    // Открываем таблицу
    var remoteSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = remoteSpreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log("  ⊗ " + operatorName + ": лист '" + sheetName + "' не найден");
      return null;
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) {
      Logger.log("  ⊗ " + operatorName + ": недостаточно данных");
      return null;
    }
    
    // Читаем первые 2 колонки (оптимизация - меньше данных)
    var data = sheet.getRange(1, 1, Math.min(lastRow, 40), 2).getValues();
    
    // Структура результата
    var stats = {
      period: "",
      updated: "",
      fixes: 0,
      unique: 0,
      open: 0,
      closed: 0,
      s1: 0, s2: 0, s3: 0, s4: 0, s5: 0
    };
    
    // Быстрый парсинг - один проход
    for (var i = 0; i < data.length; i++) {
      var a = String(data[i][0] || "").toLowerCase().trim();
      var b = data[i][1];
      var bNum = parseInt(b) || 0;
      
      // Служебная информация
      if (a.indexOf("период") > -1 && a.indexOf(":") > -1) stats.period = String(data[i][0]);
      if (a.indexOf("обновлено") > -1 && a.indexOf(":") > -1) stats.updated = String(data[i][0]);
      
      // Основные показатели
      if (a.indexOf("всего фиксаций") > -1 && a.indexOf(":") > -1) stats.fixes = bNum;
      if (a.indexOf("уникальных записей") > -1 && a.indexOf(":") > -1) stats.unique = bNum;
      
      // Открыто/Закрыто (↳ или "из них")
      if (a.indexOf("↳") > -1 || a.indexOf("из них") > -1) {
        if (a.indexOf("открыто") > -1) stats.open = bNum;
        if (a.indexOf("закрыто") > -1) stats.closed = bNum;
      }
      
      // Статусы (•)
      if (a.indexOf("•") > -1) {
        if (a.indexOf("отрицательный") > -1) stats.s1 = bNum;
        else if (a.indexOf("положительный") > -1) stats.s2 = bNum;
        else if (a.indexOf("заявка закрыта") > -1) stats.s3 = bNum;
        else if (a.indexOf("открыть карту") > -1) stats.s4 = bNum;
        else if (a.indexOf("тиббиёт") > -1 || a.indexOf("ходими") > -1) stats.s5 = bNum;
      }
    }
    
    // Валидация
    if (stats.fixes === 0 && stats.unique === 0) {
      Logger.log("  ⊗ " + operatorName + ": нет данных (фиксаций=0)");
      return null;
    }
    
    Logger.log("  ✓ " + operatorName + ": фиксаций=" + stats.fixes + ", уникальных=" + stats.unique);
    
    return [
      operatorName,
      stats.updated,
      stats.period,
      stats.fixes,
      stats.unique,
      stats.open,
      stats.closed,
      stats.s1,
      stats.s2,
      stats.s3,
      stats.s4,
      stats.s5
    ];
    
  } catch (err) {
    Logger.log("  ✗ " + operatorName + ": ОШИБКА - " + err.message);
    return null;
  }
}

/**
 * Парсит дату из формата "dd.MM.yyyy (EEE)"
 */
function parseDate(dateStr) {
  if (!dateStr) return new Date(0);
  
  if (dateStr instanceof Date) return dateStr;
  
  var str = String(dateStr).trim();
  
  // Формат: "25.11.2025 (Mon)"
  var match = str.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (match) {
    var day = parseInt(match[1], 10);
    var month = parseInt(match[2], 10) - 1;
    var year = parseInt(match[3], 10);
    return new Date(year, month, day);
  }
  
  return new Date(str);
}

// =============================================================================
// ПОМЕСЯЧНАЯ СТАТИСТИКА
// =============================================================================

/**
 * Создает отдельные листы для каждого месяца с детальной статистикой
 * ПОЭТАПНАЯ версия - обрабатывает по одному месяцу за раз
 */
function createMonthlyStatistics() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var operators = getOperatorList();
  
  // Получаем очередь месяцев
  var monthQueue = getMonthQueue();
  
  // Если очередь пустая - сканируем все месяцы
  if (monthQueue.length === 0) {
    Logger.log("Сканирование всех месяцев...");
    monthQueue = scanAllMonths(operators);
    saveMonthQueue(monthQueue);
    Logger.log("Найдено месяцев: " + monthQueue.length);
  }
  
  // Обрабатываем по одному месяцу за раз
  var processed = 0;
  for (var i = 0; i < monthQueue.length; i++) {
    if (monthQueue[i].status === "completed") {
      processed++;
      continue;
    }
    
    // Проверка таймаута (5 минут)
    var elapsed = new Date().getTime() - startTime;
    if (elapsed > MAX_EXECUTION_TIME) {
      Logger.log("⚠️ Таймаут: сохраняем прогресс по месяцам (прошло " + Math.round(elapsed/1000) + " сек)");
      saveMonthQueue(monthQueue);
      return;
    }
    
    var monthKey = monthQueue[i].month;
    Logger.log("Обработка месяца: " + monthKey);
    
    // Собираем данные для этого месяца
    var monthData = collectMonthData(operators, monthKey);
    
    // Создаем лист только если есть данные
    if (monthData && monthData.length > 0) {
      createMonthSheet(ss, monthKey, monthData);
      Logger.log("Создан лист для месяца " + monthKey + " с " + monthData.length + " записями");
    } else {
      Logger.log("Нет данных для месяца " + monthKey + ", пропускаем");
    }
    
    monthQueue[i].status = "completed";
    saveMonthQueue(monthQueue);
    processed++;
  }
  
  // Все месяцы обработаны
  if (processed === monthQueue.length) {
    clearMonthQueue();
    Logger.log("✓ Все месяцы обработаны: " + processed);
  }
  
  var endTime = new Date().getTime();
  Logger.log("createMonthlyStatistics завершено за " + Math.round((endTime - startTime) / 1000) + " сек");
}

/**
 * Сканирует все месяцы - УМНАЯ версия с анализом содержимого
 * Анализирует данные внутри листов, а не только названия
 */
function scanAllMonths(operators) {
  var monthsSet = {};
  
  Logger.log("Умное сканирование месяцев у " + operators.length + " операторов...");
  
  for (var i = 0; i < operators.length; i++) {
    var op = operators[i];
    if (op.status.toLowerCase() !== "активен") continue;
    
    try {
      var remoteSpreadsheet = SpreadsheetApp.openById(op.spreadsheetId);
      var sheets = remoteSpreadsheet.getSheets();
      
      Logger.log("  " + op.name + ": анализ " + sheets.length + " листов...");
      
      for (var j = 0; j < sheets.length; j++) {
        var sheet = sheets[j];
        var sheetName = sheet.getName();
        
        // Пропускаем служебные листы
        if (sheetName === "Статистика" || sheetName === "Предыдущий месяц" || 
            sheetName === "Сводка по дням" || sheetName === "Настройки") {
          continue;
        }
        
        // Анализируем содержимое листа
        var lastRow = sheet.getLastRow();
        if (lastRow < 2) continue;
        
        // Читаем колонку с датами (обычно столбец I, индекс 9)
        var dateColumn = sheet.getRange(2, 9, Math.min(lastRow - 1, 100), 1).getValues();
        
        for (var d = 0; d < dateColumn.length; d++) {
          var dateValue = dateColumn[d][0];
          if (!dateValue) continue;
          
          var date = parseDateFromString(dateValue);
          if (date && date.getFullYear() > 2000) {
            var monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM.yyyy");
            if (!monthsSet[monthKey]) {
              monthsSet[monthKey] = true;
              Logger.log("    Найден месяц: " + monthKey + " (лист: " + sheetName + ")");
            }
          }
        }
      }
      
    } catch (err) {
      Logger.log("  ✗ Ошибка для " + op.name + ": " + err.message);
    }
  }
  
  var monthKeys = Object.keys(monthsSet).sort().reverse();
  Logger.log("Найдено уникальных месяцев: " + monthKeys.length + " -> " + monthKeys.join(", "));
  
  return monthKeys.map(function(month) {
    return {month: month, status: "pending"};
  });
}

/**
 * Собирает данные для конкретного месяца - УМНАЯ версия с анализом содержимого
 * Ищет данные по датам внутри листов, группирует по месяцам
 */
function collectMonthData(operators, monthKey) {
  var monthData = [];
  
  Logger.log("\n▶ Сбор данных для месяца " + monthKey + "...");
  
  for (var i = 0; i < operators.length; i++) {
    var op = operators[i];
    if (op.status.toLowerCase() !== "активен") continue;
    
    try {
      var remoteSpreadsheet = SpreadsheetApp.openById(op.spreadsheetId);
      var sheets = remoteSpreadsheet.getSheets();
      var operatorStats = null;
      
      // Проходим по всем листам и ищем данные для нужного месяца
      for (var j = 0; j < sheets.length; j++) {
        var sheet = sheets[j];
        var sheetName = sheet.getName();
        
        // Пропускаем служебные листы
        if (sheetName === "Статистика" || sheetName === "Предыдущий месяц" || 
            sheetName === "Сводка по дням" || sheetName === "Настройки") {
          continue;
        }
        
        // Читаем данные из листа и фильтруем по месяцу
        var stats = readArchiveStatsForMonth(sheet, op.name, monthKey);
        if (stats) {
          // Объединяем статистику (может быть несколько листов с данными за месяц)
          if (!operatorStats) {
            operatorStats = stats;
          } else {
            // Суммируем показатели
            for (var k = 3; k < 12; k++) {
              operatorStats[k] = (operatorStats[k] || 0) + (stats[k] || 0);
            }
          }
        }
      }
      
      if (operatorStats) {
        monthData.push(operatorStats);
        Logger.log("  ✓ " + op.name + ": фиксаций=" + operatorStats[3]);
      } else {
        Logger.log("  ⊗ " + op.name + ": нет данных за " + monthKey);
      }

    } catch (err) {
      Logger.log("  ✗ " + op.name + ": " + err.message);
    }
  }
  
  Logger.log("Итого: " + monthData.length + " операторов за " + monthKey + "\n");
  return monthData;
}

/**
 * Получает очередь месяцев
 */
function getMonthQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }
  
  // Ищем секцию месяцев (после разделителя "Месяцы:")
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var queue = [];
  var inMonthSection = false;
  
  for (var i = 0; i < data.length; i++) {
    var cellValue = String(data[i][0] || "").trim();
    
    // Нашли разделитель - следующие строки это месяцы
    if (cellValue === "Месяцы:") {
      inMonthSection = true;
      continue;
    }
    
    // Читаем только после разделителя
    if (inMonthSection && cellValue) {
      // Проверяем формат MM.yyyy
      if (cellValue.match(/\d{2}\.\d{4}/)) {
        queue.push({
          month: cellValue,
          status: String(data[i][1] || "pending").trim()
        });
      }
    }
  }
  
  Logger.log("Прочитано месяцев из очереди: " + queue.length);
  return queue;
}

/**
 * Сохраняет очередь месяцев
 */
function saveMonthQueue(monthQueue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_SHEET_NAME);
    sheet.hideSheet();
  }
  
  // Читаем текущую основную очередь (НЕ через getTaskQueue, чтобы не потерять данные)
  var existingData = [];
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var taskName = String(data[i][0] || "").trim();
      if (taskName === "Месяцы:" || taskName === "") break;
      if (taskName) {
        existingData.push([data[i][0], data[i][1]]);
      }
    }
  }
  
  // Очищаем и пишем заново
  sheet.clear();
  sheet.getRange("A1:B1").setValues([["Задача", "Статус"]]);
  
  var row = 2;
  
  // Основная очередь (сохраняем как есть)
  if (existingData.length > 0) {
    sheet.getRange(row, 1, existingData.length, 2).setValues(existingData);
    row += existingData.length;
  }
  
  // Разделитель
  sheet.getRange(row, 1).setValue("Месяцы:");
  row++;
  
  // Очередь месяцев
  if (monthQueue && monthQueue.length > 0) {
    var monthData = monthQueue.map(function(item) {
      return [item.month, item.status];
    });
    sheet.getRange(row, 1, monthData.length, 2).setValues(monthData);
    Logger.log("Сохранено месяцев: " + monthQueue.length);
  }
}

// Функция clearTaskQueue удалена - больше не нужна

/**
 * Очищает очередь месяцев
 */
function clearMonthQueue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  
  if (!sheet) return;
  
  // Удаляем только секцию месяцев
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var deleteFrom = -1;
  
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === "Месяцы:") {
      deleteFrom = i + 2; // +2 т.к. строка 1 - заголовок, строка i+2 - "Месяцы:"
      break;
    }
  }
  
  if (deleteFrom > 0 && deleteFrom <= sheet.getLastRow()) {
    sheet.getRange(deleteFrom, 1, sheet.getLastRow() - deleteFrom + 1, 2).clearContent();
  }
}

/**
 * Создает или обновляет лист для конкретного месяца
 */
function createMonthSheet(ss, monthKey, data) {
  // Проверяем наличие данных
  if (!data || data.length === 0) {
    Logger.log("Нет данных для создания листа месяца " + monthKey);
    return;
  }
  
  var sheetName = MONTHLY_STATS_PREFIX + monthKey;
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  sheet.clear();
  
  // Заголовки
  var headers = [
    "ФИО оператора",
    "Дата обновления",
    "Период",
    "Всего фиксаций",
    "Уникальных записей",
    "Открыто",
    "Закрыто",
    "Отрицательный",
    "Положительный",
    "Заявка закрыта",
    "Открыть карту",
    "Тиббиёт ходими"
  ];
  
  var allData = [headers];
  
  // Добавляем данные
  for (var i = 0; i < data.length; i++) {
    // Проверяем что каждая строка имеет правильное количество столбцов
    if (data[i] && data[i].length === headers.length) {
      allData.push(data[i]);
    } else {
      Logger.log("Пропущена некорректная строка данных для " + monthKey);
    }
  }

  // Добавляем итоговую строку только если есть валидные данные
  if (allData.length > 1) {
    var totals = calculateTotals(allData.slice(1)); // Берем без заголовка
    allData.push([]);
    allData.push(totals);
  }

  // Записываем данные только если есть хотя бы одна строка данных (кроме заголовка)
  if (allData.length > 2) {
    sheet.getRange(1, 1, allData.length, headers.length).setValues(allData);

    // Форматирование заголовка
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("#ffffff");

    // Форматирование итоговой строки
    var totalRow = allData.length;
    sheet.getRange(totalRow, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#f3f3f3");

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);

    Logger.log("Создан лист: " + sheetName + " с " + (allData.length - 1) + " строками");
  } else {
    Logger.log("Нет валидных данных для листа " + sheetName + ", лист не создан");
  }
}

/**
 * Читает статистику с архивного листа для конкретного месяца
 * Фильтрует данные по месяцу из колонки с датами
 */
function readArchiveStatsForMonth(sheet, operatorName, targetMonth) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    var data = sheet.getRange(2, 2, lastRow - 1, 8).getValues(); // B-I колонки
    
    var totalFixes = 0;
    var uniqueCards = {};
    var closedCards = {};
    var openCards = {};
    var statusCount = {
      "отрицательный": {},
      "положительный": {},
      "заявка закрыта (не удалось дозвониться)": {},
      "открыть карту": {},
      "тиббиёт ходими аризаси": {}
    };

    var CLOSED_STATUSES = [
      "отрицательный",
      "положительный",
      "заявка закрыта (не удалось дозвониться)",
      "открыть карту",
      "тиббиёт ходими аризаси"
    ];

    var minDate = null;
    var maxDate = null;
    var foundRecords = false;

    for (var i = 0; i < data.length; i++) {
      var cardNum = String(data[i][0] || "").trim(); // B - номер карты
      var status = String(data[i][3] || "").trim().toLowerCase(); // E - статус
      var dateStr = data[i][7]; // I - дата

      if (!cardNum || !dateStr) continue;

      var date = parseDateFromString(dateStr);
      if (!date || date.getFullYear() < 2000) continue;
      
      // Проверяем, принадлежит ли дата нужному месяцу
      var recordMonth = Utilities.formatDate(date, Session.getScriptTimeZone(), "MM.yyyy");
      if (recordMonth !== targetMonth) continue;
      
      foundRecords = true;
      totalFixes++;
      uniqueCards[cardNum] = true;

      if (!minDate || date < minDate) minDate = date;
      if (!maxDate || date > maxDate) maxDate = date;

      // Подсчитываем статусы
      var isClosed = false;
      for (var s = 0; s < CLOSED_STATUSES.length; s++) {
        if (status === CLOSED_STATUSES[s].toLowerCase()) {
          closedCards[cardNum] = true;
          statusCount[CLOSED_STATUSES[s]][cardNum] = true;
          isClosed = true;
          break;
        }
      }
      
      if (!isClosed) {
        openCards[cardNum] = true;
      }
    }
    
    if (!foundRecords) return null;
    
    var period = "";
    if (minDate && maxDate) {
      var tz = Session.getScriptTimeZone();
      period = "Период: " + 
               Utilities.formatDate(minDate, tz, "dd.MM.yyyy") + 
               " - " + 
               Utilities.formatDate(maxDate, tz, "dd.MM.yyyy");
    }
    
    return [
      operatorName,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss"),
      period,
      totalFixes,
      Object.keys(uniqueCards).length,
      Object.keys(openCards).length,
      Object.keys(closedCards).length,
      Object.keys(statusCount["отрицательный"]).length,
      Object.keys(statusCount["положительный"]).length,
      Object.keys(statusCount["заявка закрыта (не удалось дозвониться)"]).length,
      Object.keys(statusCount["открыть карту"]).length,
      Object.keys(statusCount["тиббиёт ходими аризаси"]).length
    ];

  } catch (err) {
    Logger.log("    Ошибка чтения " + sheet.getName() + ": " + err.message);
    return null;
  }
}

/**
 * Читает статистику с архивного листа (старая версия, оставлена для совместимости)
 * ОПТИМИЗИРОВАНО: читает только нужные столбцы
 */
function readArchiveStats(sheet, operatorName) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("Лист пустой или недостаточно данных: " + sheet.getName());
      return null;
    }

    var data = sheet.getRange(2, 2, lastRow - 1, 8).getValues(); // B-I колонки
    Logger.log("Обработка архивного листа " + sheet.getName() + ": " + data.length + " строк");

    var totalFixes = 0;
    var uniqueCards = {};
    var closedCards = {};
    var openCards = {};
    var statusCount = {
      "отрицательный": {},
      "положительный": {},
      "заявка закрыта (не удалось дозвониться)": {},
      "открыть карту": {},
      "тиббиёт ходими аризаси": {}
    };

    var CLOSED_STATUSES = [
      "отрицательный",
      "положительный",
      "заявка закрыта (не удалось дозвониться)",
      "открыть карту",
      "тиббиёт ходими аризаси"
    ];

    var period = "";
    var minDate = null;
    var maxDate = null;

    for (var i = 0; i < data.length; i++) {
      var cardNum = String(data[i][0] || "").trim(); // B - номер карты
      var status = String(data[i][3] || "").trim().toLowerCase(); // E - статус
      var dateStr = data[i][7]; // I - дата

      if (!cardNum) continue;

      totalFixes++;
      uniqueCards[cardNum] = true;

      // Определяем период
      if (dateStr) {
        var date = parseDateFromString(dateStr);
        if (date) {
          if (!minDate || date < minDate) minDate = date;
          if (!maxDate || date > maxDate) maxDate = date;
        }
      }

      // Подсчитываем статусы
      var isClosed = false;
      for (var s = 0; s < CLOSED_STATUSES.length; s++) {
        if (status === CLOSED_STATUSES[s].toLowerCase()) {
          closedCards[cardNum] = true;
          statusCount[CLOSED_STATUSES[s]][cardNum] = true;
          isClosed = true;
          break;
        }
      }
      
      if (!isClosed) {
        openCards[cardNum] = true;
      }
    }
    
    // Формируем период
    if (minDate && maxDate) {
      var tz = Session.getScriptTimeZone();
      period = "Период: " + 
               Utilities.formatDate(minDate, tz, "dd.MM.yyyy") + 
               " - " + 
               Utilities.formatDate(maxDate, tz, "dd.MM.yyyy");
    }
    
    var result = [
      operatorName,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss"),
      period,
      totalFixes,
      Object.keys(uniqueCards).length,
      Object.keys(openCards).length,
      Object.keys(closedCards).length,
      Object.keys(statusCount["отрицательный"]).length,
      Object.keys(statusCount["положительный"]).length,
      Object.keys(statusCount["заявка закрыта (не удалось дозвониться)"]).length,
      Object.keys(statusCount["открыть карту"]).length,
      Object.keys(statusCount["тиббиёт ходими аризаси"]).length
    ];
    
    Logger.log("Результат обработки листа " + sheet.getName() + ": " + JSON.stringify(result));
    return result;

  } catch (err) {
    Logger.log("Ошибка чтения архивного листа: " + err);
    return null;
  }
}

/**
 * Извлекает месяц из строки периода
 */
function extractMonthFromPeriod(periodStr) {
  if (!periodStr) return null;
  
  // Ищем дату в формате dd.MM.yyyy
  var match = periodStr.match(/(\d{2})\.(\d{2})\.(\d{4})/);
  if (match) {
    return match[2] + "." + match[3]; // MM.yyyy
  }
  
  return null;
}

/**
 * Парсит дату из строки или объекта Date - УЛУЧШЕННАЯ версия
 * Поддерживает множество форматов
 */
function parseDateFromString(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  
  var str = String(value).trim();
  
  // Формат: "01.12.2024 10:30:45"
  var pattern1 = /(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})/;
  var match1 = str.match(pattern1);
  if (match1) {
    return new Date(
      parseInt(match1[3], 10),
      parseInt(match1[2], 10) - 1,
      parseInt(match1[1], 10),
      parseInt(match1[4], 10),
      parseInt(match1[5], 10),
      parseInt(match1[6], 10)
    );
  }
  
  // Формат: "01.12.2024"
  var pattern2 = /(\d{2})\.(\d{2})\.(\d{4})/;
  var match2 = str.match(pattern2);
  if (match2) {
    return new Date(
      parseInt(match2[3], 10),
      parseInt(match2[2], 10) - 1,
      parseInt(match2[1], 10)
    );
  }
  
  // Пробуем стандартный парсинг
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  return null;
}

/**
 * Вычисляет итоговые суммы по всем операторам
 */
function calculateTotals(data) {
  var totals = [
    "ИТОГО:",
    "",
    "",
    0, // Всего фиксаций
    0, // Уникальных записей
    0, // Открыто
    0, // Закрыто
    0, // Отрицательный
    0, // Положительный
    0, // Заявка закрыта
    0, // Открыть карту
    0  // Тиббиёт ходими
  ];
  
  for (var i = 0; i < data.length; i++) {
    for (var j = 3; j < 12; j++) {
      var val = parseInt(data[i][j]) || 0;
      totals[j] += val;
    }
  }
  
  return totals;
}

// =============================================================================
// КОНЕЦ СКРИПТА СБОРЩИКА
// =============================================================================

// ФУНКЦИИ ОЧЕРЕДЕЙ БОЛЬШЕ НЕ НУЖНЫ - КОД УПРОЩЁН

// ============================================
// ЮРА ТРАНСПОРТЕЙШН — CRM ARCHIVE SYSTEM v1.0
// Apps Script API для архівування записів CRM
// Таблиця: Архіви (ID: 1Kmf6NF1sJUi-j3SamrhUqz337pcZSvZCUkGxBzari6U)
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Архіви" → Розширення → Apps Script
// 2. Видали весь старий код і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в CRM HTML файл замість YOUR_ARCHIVE_API_URL_HERE
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

// ID архівної таблиці
var ARCHIVE_SS_ID = '1Kmf6NF1sJUi-j3SamrhUqz337pcZSvZCUkGxBzari6U';

// Архівні аркуші (в таблиці "Архіви")
var ARCHIVE_SHEETS = {
  POSYLKY: 'Посилки',
  POSYLKY_ROUTE: 'Посилки маршрут',
  PASAZHYRY: 'Пасажири',
  PASAZHYRY_ROUTE: 'Пасажири маршрут'
};

// Аркуш логів (в таблиці "Архіви")
var LOG_SHEET_NAME = 'Логи';

// ============================================
// ДЖЕРЕЛА ДАНИХ — звідки архівуємо
// ============================================
var SOURCES = {
  BOT_POSYLKY: {
    id: '1E9wYOmVTtlDc52kQAekSpc6rw7Mdnot-m24pRvTUlaY',
    name: 'Logistics-Cargo',
    type: 'posylky',
    sheets: ['Реєстрація ТТН', 'Виклик курєра', 'РОБОЧА'],
    archiveSheet: 'Посилки',
    // Структура: A(0)-T(19)=дані, U(20)=Статус, V(21)=Дата архів, W(22)=ARCHIVE_ID
    dataCols: 21,       // кількість колонок даних (A-U, 0-20)
    totalCols: 23,      // всього колонок (A-W, 0-22)
    archiveIdCol: 22,   // W — ARCHIVE_ID в джерелі
    statusCol: 20,      // U — Статус
    dateArchiveCol: 21, // V — Дата архів
    idCol: 13           // N — ІД запису
  },
  ROUTE_POSYLKY: {
    id: '17g3TFYg11EqdQ9eGrOKQV3n_nqPDFx7dqsJVaGWeDOo',
    name: 'Маршрути Посилки',
    type: 'posylky',
    sheets: ['Братислава марш.', 'Словаччина марш.', 'Нітра марш.', 'Кошице+прешов марш.'],
    archiveSheet: 'Посилки маршрут',
    // Структура: A(0)-T(19)=дані, U(20)=Статус, V(21)-Z(25)=архівні поля
    dataCols: 21,
    totalCols: 26,
    archiveIdCol: 25,   // Z — ARCHIVE_ID
    statusCol: 20,      // U — Статус
    dateArchiveCol: 21, // V — DATE_ARCHIVE
    idCol: 13           // N — ІД
  },
  BOT_PASAZHYRY: {
    id: '1SvWAVYNKkWl7Sx_wWhTWPlDyKU9hIQCQzcWBTUGG2i8',
    name: 'Бот Пасажири',
    type: 'pasazhyry',
    sheets: ['Україна-єв', 'Європа-ук', 'РОБОЧИЙ'],
    archiveSheet: 'Пасажири',
    // Структура: A(0)-Q(16)=дані, R(17)=Статус, S(18)=Дата архів, T(19)=ARCHIVE_ID
    dataCols: 18,
    totalCols: 20,
    archiveIdCol: 19,   // T — ARCHIVE_ID
    statusCol: 17,      // R — Статус
    dateArchiveCol: 18, // S — Дата архів
    idCol: 10           // K — ІД
  },
  ROUTE_PASAZHYRY: {
    id: '1fYO1ClIP26S4xYgcsT_0LVCWVrqkAL5MkehXvL-Yni0',
    name: 'Маршрут Пасажири',
    type: 'pasazhyry',
    sheets: ['Пас. Маршрут 1', 'Пас. Маршрут 2', 'Пас. Маршрут 3'],
    archiveSheet: 'Пасажири маршрут',
    // Структура: A(0)-Q(16)=дані, R(17)=Статус, S(18)-W(22)=архівні поля
    dataCols: 18,
    totalCols: 23,
    archiveIdCol: 22,   // W — ARCHIVE_ID
    statusCol: 17,      // R — Статус
    dateArchiveCol: 18, // S — DATE_ARCHIVE
    idCol: 10           // K — ІД
  }
};

// Кількість колонок в архівних аркушах
var ARCHIVE_TOTAL_COLS = {
  posylky: 26,      // A-Z (26 колонок)
  pasazhyry: 23     // A-W (23 колонки)
};

// Індекси архівних метаданих в архівній таблиці
var ARCHIVE_META = {
  posylky: {
    // Архів "Посилки" / "Посилки маршрут": 26 cols
    DATE_ARCHIVE: 21,   // V
    ARCHIVED_BY: 22,    // W
    ARCHIVE_REASON: 23, // X
    SOURCE_SHEET: 24,   // Y
    ARCHIVE_ID: 25      // Z
  },
  pasazhyry: {
    // Архів "Пасажири" / "Пасажири маршрут": 23 cols
    DATE_ARCHIVE: 18,   // S
    ARCHIVED_BY: 19,    // T
    ARCHIVE_REASON: 20, // U
    SOURCE_SHEET: 21,   // V
    ARCHIVE_ID: 22      // W
  }
};

// Зворотній маппінг: архівний аркуш → source key
var ARCHIVE_TO_SOURCE = {
  'Посилки': 'BOT_POSYLKY',
  'Посилки маршрут': 'ROUTE_POSYLKY',
  'Пасажири': 'BOT_PASAZHYRY',
  'Пасажири маршрут': 'ROUTE_PASAZHYRY'
};

// Статуси для автоархівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred', 'completed', 'cancelled'];

// ============================================
// ГОЛОВНИЙ ОБРОБНИК — doPost
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    switch (action) {
      // --- АРХІВАЦІЯ ---
      case 'archiveByStatus':
        return respond(archiveByStatus(data));

      case 'archiveByIds':
        return respond(archiveByIds(data));

      case 'archiveByAge':
        return respond(archiveByAge(data));

      case 'archiveAll':
        return respond(archiveAll(data));

      // --- ПЕРЕГЛЯД ---
      case 'getArchived':
        return respond(getArchived(data));

      case 'searchArchive':
        return respond(searchArchive(data));

      case 'getStats':
        return respond(getStats(data));

      // --- ВІДНОВЛЕННЯ ---
      case 'restoreRecord':
        return respond(restoreRecord(data));

      // --- ВИДАЛЕННЯ З АРХІВУ ---
      case 'deleteFromArchive':
        return respond(deleteFromArchive(data));

      default:
        return respond({ success: false, error: 'Невідома дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// doGet — для перевірки здоров'я та статистики
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'health';

    switch (action) {
      case 'health':
        return respond({
          success: true,
          version: '1.0',
          service: 'CRM Archive — ЮРА ТРАНСПОРТЕЙШН',
          timestamp: new Date().toISOString()
        });

      case 'getStats':
        return respond(getStats({}));

      default:
        return respond({ success: false, error: 'Невідома GET дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// АРХІВАЦІЯ ЗА СТАТУСОМ
// Архівує всі записи з вказаними статусами
// ============================================
// Параметри:
//   source: 'BOT_POSYLKY' | 'ROUTE_POSYLKY' | 'BOT_PASAZHYRY' | 'ROUTE_PASAZHYRY' | 'ALL'
//   statuses: ['archived', 'refused', ...] (необов'язково, за замовч. ARCHIVE_STATUSES)
//   user: ім'я користувача
//   reason: причина архівації
//   deleteFromSource: true/false (чи видаляти з джерела, за замовч. true)
// ============================================
function archiveByStatus(data) {
  var sourceKey = data.source || 'ALL';
  var statuses = data.statuses || ARCHIVE_STATUSES;
  var user = data.user || 'system';
  var reason = data.reason || 'status';
  var deleteFromSource = data.deleteFromSource !== false;

  // Нормалізуємо статуси до нижнього регістру
  for (var i = 0; i < statuses.length; i++) {
    statuses[i] = statuses[i].toLowerCase().trim();
  }

  var totalArchived = 0;
  var results = [];

  // Визначаємо які джерела обробляти
  var sourceKeys = [];
  if (sourceKey === 'ALL') {
    sourceKeys = ['BOT_POSYLKY', 'ROUTE_POSYLKY', 'BOT_PASAZHYRY', 'ROUTE_PASAZHYRY'];
  } else {
    sourceKeys = [sourceKey];
  }

  for (var s = 0; s < sourceKeys.length; s++) {
    var config = SOURCES[sourceKeys[s]];
    if (!config) {
      results.push({ source: sourceKeys[s], error: 'Невідоме джерело' });
      continue;
    }

    var sourceResult = archiveFromSource(config, function(row) {
      var status = String(row[config.statusCol] || '').toLowerCase().trim();
      return statuses.indexOf(status) !== -1;
    }, user, reason, deleteFromSource);

    totalArchived += sourceResult.archived;
    results.push({
      source: sourceKeys[s],
      name: config.name,
      archived: sourceResult.archived,
      sheets: sourceResult.sheets
    });
  }

  return {
    success: true,
    totalArchived: totalArchived,
    results: results,
    timestamp: getNow()
  };
}

// ============================================
// АРХІВАЦІЯ ЗА ІД
// Архівує конкретні записи за їхнім ІД
// ============================================
// Параметри:
//   source: ключ джерела
//   sheet: назва аркуша в джерелі (необов'язково — шукає по всіх)
//   ids: масив ІД записів для архівації
//   user, reason, deleteFromSource
// ============================================
function archiveByIds(data) {
  var sourceKey = data.source;
  var targetSheet = data.sheet || null;
  var ids = data.ids || [];
  var user = data.user || 'system';
  var reason = data.reason || 'manual';
  var deleteFromSource = data.deleteFromSource !== false;

  if (!sourceKey || !ids.length) {
    return { success: false, error: 'Потрібні source та ids' };
  }

  var config = SOURCES[sourceKey];
  if (!config) {
    return { success: false, error: 'Невідоме джерело: ' + sourceKey };
  }

  // Нормалізуємо IDs до рядків
  var idSet = {};
  for (var i = 0; i < ids.length; i++) {
    idSet[String(ids[i]).trim()] = true;
  }

  var sourceResult = archiveFromSource(config, function(row) {
    var recordId = String(row[config.idCol] || '').trim();
    return recordId && idSet[recordId];
  }, user, reason, deleteFromSource, targetSheet);

  return {
    success: true,
    totalArchived: sourceResult.archived,
    requestedIds: ids.length,
    sheets: sourceResult.sheets,
    timestamp: getNow()
  };
}

// ============================================
// АРХІВАЦІЯ ЗА ВІКОМ
// Архівує записи старші за X днів
// ============================================
// Параметри:
//   source: ключ джерела або 'ALL'
//   days: кількість днів (за замовч. 30)
//   dateColumn: індекс колонки з датою (необов'язково)
//   user, reason, deleteFromSource
// ============================================
function archiveByAge(data) {
  var sourceKey = data.source || 'ALL';
  var days = data.days || 30;
  var user = data.user || 'auto';
  var reason = data.reason || 'age_' + days + '_days';
  var deleteFromSource = data.deleteFromSource !== false;

  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - days);
  var cutoffTime = cutoffDate.getTime();

  var totalArchived = 0;
  var results = [];

  var sourceKeys = [];
  if (sourceKey === 'ALL') {
    sourceKeys = ['BOT_POSYLKY', 'ROUTE_POSYLKY', 'BOT_PASAZHYRY', 'ROUTE_PASAZHYRY'];
  } else {
    sourceKeys = [sourceKey];
  }

  for (var s = 0; s < sourceKeys.length; s++) {
    var config = SOURCES[sourceKeys[s]];
    if (!config) continue;

    // Колонка дати: для посилок це P(15)="дата оформлення", для пасажирів це P(15) теж
    var dateCol = data.dateColumn || 15;

    var sourceResult = archiveFromSource(config, function(row) {
      var dateVal = row[dateCol];
      if (!dateVal) return false;

      var rowDate = parseDate(dateVal);
      if (!rowDate) return false;

      return rowDate.getTime() < cutoffTime;
    }, user, reason, deleteFromSource);

    totalArchived += sourceResult.archived;
    results.push({
      source: sourceKeys[s],
      name: config.name,
      archived: sourceResult.archived
    });
  }

  return {
    success: true,
    totalArchived: totalArchived,
    days: days,
    results: results,
    timestamp: getNow()
  };
}

// ============================================
// АРХІВАЦІЯ ВСІХ ЗАПИСІВ З ДЖЕРЕЛА
// (для повного очищення маршрутних аркушів)
// ============================================
// Параметри:
//   source: ключ джерела
//   sheet: конкретний аркуш (необов'язково)
//   user, reason, deleteFromSource
// ============================================
function archiveAll(data) {
  var sourceKey = data.source;
  var targetSheet = data.sheet || null;
  var user = data.user || 'system';
  var reason = data.reason || 'full_archive';
  var deleteFromSource = data.deleteFromSource !== false;

  if (!sourceKey) {
    return { success: false, error: 'Потрібен source' };
  }

  var config = SOURCES[sourceKey];
  if (!config) {
    return { success: false, error: 'Невідоме джерело: ' + sourceKey };
  }

  // Архівуємо все (фільтр завжди true, але пропускаємо порожні)
  var sourceResult = archiveFromSource(config, function(row) {
    return !isEmptyRow(row, config.dataCols);
  }, user, reason, deleteFromSource, targetSheet);

  return {
    success: true,
    totalArchived: sourceResult.archived,
    sheets: sourceResult.sheets,
    timestamp: getNow()
  };
}

// ============================================
// ОСНОВНА ФУНКЦІЯ АРХІВАЦІЇ
// Читає джерело, фільтрує, пише в архів, видаляє з джерела
// ============================================
function archiveFromSource(config, filterFn, user, reason, deleteFromSource, targetSheet) {
  var sourceSS;
  try {
    sourceSS = SpreadsheetApp.openById(config.id);
  } catch (err) {
    return { archived: 0, sheets: [], error: 'Не вдалося відкрити джерело: ' + err.toString() };
  }

  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var archiveSheet = archiveSS.getSheetByName(config.archiveSheet);
  if (!archiveSheet) {
    return { archived: 0, sheets: [], error: 'Архівний аркуш не знайдено: ' + config.archiveSheet };
  }

  var meta = ARCHIVE_META[config.type];
  var archiveColCount = ARCHIVE_TOTAL_COLS[config.type];
  var totalArchived = 0;
  var sheetResults = [];

  // Визначаємо які аркуші обробляти
  var sheetsToProcess = targetSheet ? [targetSheet] : config.sheets;

  for (var s = 0; s < sheetsToProcess.length; s++) {
    var sheetName = sheetsToProcess[s];
    var sheet = sourceSS.getSheetByName(sheetName);

    if (!sheet) {
      sheetResults.push({ sheet: sheetName, archived: 0, error: 'Не знайдено' });
      continue;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      sheetResults.push({ sheet: sheetName, archived: 0, note: 'Порожній' });
      continue;
    }

    // Читаємо всі дані одним запитом
    var dataRange = sheet.getRange(2, 1, lastRow - 1, config.totalCols);
    var allValues = dataRange.getValues();

    // Фільтруємо рядки для архівації
    var rowsToArchive = []; // { rowIndex (0-based in allValues), rowNum (1-based in sheet) }
    var archiveRows = [];   // готові рядки для архіву

    for (var i = 0; i < allValues.length; i++) {
      var row = allValues[i];

      // Пропускаємо порожні рядки
      if (isEmptyRow(row, config.dataCols)) continue;

      // Пропускаємо вже заархівовані (мають ARCHIVE_ID)
      var existingArchiveId = String(row[config.archiveIdCol] || '').trim();
      if (existingArchiveId) continue;

      // Застосовуємо фільтр
      if (!filterFn(row)) continue;

      // Будуємо рядок для архіву
      var archiveId = generateArchiveId();
      var archiveRow = buildArchiveRow(row, config, meta, archiveColCount, sheetName, user, reason, archiveId);

      rowsToArchive.push({
        index: i,
        rowNum: i + 2, // рядок в таблиці (1-based, +1 за заголовок)
        archiveId: archiveId
      });
      archiveRows.push(archiveRow);
    }

    if (archiveRows.length === 0) {
      sheetResults.push({ sheet: sheetName, archived: 0 });
      continue;
    }

    // --- КРОК 1: Записуємо в архів (batch) ---
    var archiveStartRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(archiveStartRow, 1, archiveRows.length, archiveColCount)
      .setValues(archiveRows);

    // --- КРОК 2: Записуємо ARCHIVE_ID назад у джерело ---
    for (var j = 0; j < rowsToArchive.length; j++) {
      var item = rowsToArchive[j];
      sheet.getRange(item.rowNum, config.archiveIdCol + 1).setValue(item.archiveId);

      // Ставимо дату архіву в джерелі (якщо є колонка)
      if (config.dateArchiveCol !== undefined) {
        sheet.getRange(item.rowNum, config.dateArchiveCol + 1).setValue(getNow());
      }
    }

    // --- КРОК 3: Видаляємо з джерела (знизу вгору) ---
    if (deleteFromSource) {
      // Сортуємо рядки знизу вгору для безпечного видалення
      var rowNums = [];
      for (var k = 0; k < rowsToArchive.length; k++) {
        rowNums.push(rowsToArchive[k].rowNum);
      }
      rowNums.sort(function(a, b) { return b - a; }); // від більшого до меншого

      for (var d = 0; d < rowNums.length; d++) {
        sheet.deleteRow(rowNums[d]);
      }
    }

    // --- КРОК 4: Логуємо ---
    for (var l = 0; l < rowsToArchive.length; l++) {
      var recordId = String(allValues[rowsToArchive[l].index][config.idCol] || '');
      writeArchiveLog(
        config.name,
        user,
        'ARCHIVE',
        recordId,
        sheetName,
        reason + (deleteFromSource ? ' (видалено з джерела)' : ' (залишено в джерелі)'),
        'success',
        rowsToArchive[l].archiveId
      );
    }

    totalArchived += archiveRows.length;
    sheetResults.push({ sheet: sheetName, archived: archiveRows.length });
  }

  return { archived: totalArchived, sheets: sheetResults };
}

// ============================================
// ПОБУДОВА РЯДКА АРХІВУ
// ============================================
function buildArchiveRow(sourceRow, config, meta, archiveColCount, sheetName, user, reason, archiveId) {
  var archiveRow = new Array(archiveColCount);
  for (var i = 0; i < archiveColCount; i++) {
    archiveRow[i] = '';
  }

  // Копіюємо колонки даних (включаючи Статус)
  for (var j = 0; j < config.dataCols; j++) {
    archiveRow[j] = sourceRow[j] !== undefined ? sourceRow[j] : '';
  }

  // Заповнюємо архівні метадані
  archiveRow[meta.DATE_ARCHIVE] = getNow();
  archiveRow[meta.ARCHIVED_BY] = user;
  archiveRow[meta.ARCHIVE_REASON] = reason;
  archiveRow[meta.SOURCE_SHEET] = sheetName;
  archiveRow[meta.ARCHIVE_ID] = archiveId;

  return archiveRow;
}

// ============================================
// ОТРИМАТИ АРХІВОВАНІ ЗАПИСИ
// ============================================
// Параметри:
//   archiveSheet: 'Посилки' | 'Посилки маршрут' | 'Пасажири' | 'Пасажири маршрут'
//   limit: макс. кількість записів (за замовч. 100)
//   offset: зміщення (для пагінації, за замовч. 0)
//   sortDesc: true/false — нові першими (за замовч. true)
// ============================================
function getArchived(data) {
  var archiveSheetName = data.archiveSheet;
  var limit = data.limit || 100;
  var offset = data.offset || 0;
  var sortDesc = data.sortDesc !== false;

  if (!archiveSheetName) {
    return { success: false, error: 'Потрібен archiveSheet' };
  }

  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var sheet = archiveSS.getSheetByName(archiveSheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + archiveSheetName };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, records: [], total: 0 };
  }

  var totalCols = sheet.getLastColumn();
  var allValues = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();
  var headers = sheet.getRange(1, 1, 1, totalCols).getValues()[0];

  // Фільтруємо порожні рядки
  var records = [];
  for (var i = 0; i < allValues.length; i++) {
    if (isEmptyRow(allValues[i], totalCols)) continue;

    var record = { rowNum: i + 2 };
    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c] || '').trim();
      if (key) {
        var val = allValues[i][c];
        record[key] = (val instanceof Date) ? formatDateValue(val) : String(val || '');
      }
    }
    records.push(record);
  }

  // Сортування
  if (sortDesc) {
    records.reverse();
  }

  // Пагінація
  var total = records.length;
  var paged = records.slice(offset, offset + limit);

  return {
    success: true,
    records: paged,
    total: total,
    limit: limit,
    offset: offset,
    archiveSheet: archiveSheetName
  };
}

// ============================================
// ПОШУК В АРХІВІ
// ============================================
// Параметри:
//   query: текст для пошуку
//   archiveSheet: в якому аркуші шукати (необов'язково — шукає по всіх)
//   limit: макс. результатів (за замовч. 50)
// ============================================
function searchArchive(data) {
  var query = String(data.query || '').toLowerCase().trim();
  var targetSheet = data.archiveSheet || null;
  var limit = data.limit || 50;

  if (!query) {
    return { success: false, error: 'Потрібен query' };
  }

  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var sheetsToSearch = [];

  if (targetSheet) {
    sheetsToSearch = [targetSheet];
  } else {
    sheetsToSearch = [
      ARCHIVE_SHEETS.POSYLKY,
      ARCHIVE_SHEETS.POSYLKY_ROUTE,
      ARCHIVE_SHEETS.PASAZHYRY,
      ARCHIVE_SHEETS.PASAZHYRY_ROUTE
    ];
  }

  var results = [];

  for (var s = 0; s < sheetsToSearch.length; s++) {
    var sheetName = sheetsToSearch[s];
    var sheet = archiveSS.getSheetByName(sheetName);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) continue;

    var totalCols = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, totalCols).getValues()[0];
    var allValues = sheet.getRange(2, 1, lastRow - 1, totalCols).getValues();

    for (var i = 0; i < allValues.length; i++) {
      if (results.length >= limit) break;

      var row = allValues[i];
      var rowStr = row.join(' ').toLowerCase();

      if (rowStr.indexOf(query) !== -1) {
        var record = { archiveSheet: sheetName, rowNum: i + 2 };
        for (var c = 0; c < headers.length; c++) {
          var key = String(headers[c] || '').trim();
          if (key) {
            var val = row[c];
            record[key] = (val instanceof Date) ? formatDateValue(val) : String(val || '');
          }
        }
        results.push(record);
      }
    }

    if (results.length >= limit) break;
  }

  return {
    success: true,
    results: results,
    count: results.length,
    query: query,
    timestamp: getNow()
  };
}

// ============================================
// СТАТИСТИКА АРХІВУ
// ============================================
function getStats(data) {
  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var stats = {
    total: 0,
    sheets: {}
  };

  var sheetNames = [
    ARCHIVE_SHEETS.POSYLKY,
    ARCHIVE_SHEETS.POSYLKY_ROUTE,
    ARCHIVE_SHEETS.PASAZHYRY,
    ARCHIVE_SHEETS.PASAZHYRY_ROUTE
  ];

  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var sheet = archiveSS.getSheetByName(name);

    if (!sheet) {
      stats.sheets[name] = { count: 0, exists: false };
      continue;
    }

    var count = Math.max(0, sheet.getLastRow() - 1);
    stats.sheets[name] = {
      count: count,
      exists: true,
      lastColumn: sheet.getLastColumn()
    };
    stats.total += count;
  }

  // Логи
  var logSheet = archiveSS.getSheetByName(LOG_SHEET_NAME);
  stats.logs = logSheet ? Math.max(0, logSheet.getLastRow() - 1) : 0;

  // Джерела — кількість записів для архівації
  stats.sources = {};
  for (var key in SOURCES) {
    if (!SOURCES.hasOwnProperty(key)) continue;
    var config = SOURCES[key];

    try {
      var ss = SpreadsheetApp.openById(config.id);
      var sourceStats = { total: 0, readyToArchive: 0, sheets: {} };

      for (var s = 0; s < config.sheets.length; s++) {
        var sheetName = config.sheets[s];
        var srcSheet = ss.getSheetByName(sheetName);
        if (!srcSheet) continue;

        var lastRow = srcSheet.getLastRow();
        if (lastRow < 2) {
          sourceStats.sheets[sheetName] = { total: 0, readyToArchive: 0 };
          continue;
        }

        var values = srcSheet.getRange(2, 1, lastRow - 1, config.totalCols).getValues();
        var total = 0;
        var ready = 0;

        for (var r = 0; r < values.length; r++) {
          if (isEmptyRow(values[r], config.dataCols)) continue;
          total++;

          var status = String(values[r][config.statusCol] || '').toLowerCase().trim();
          var hasArchiveId = String(values[r][config.archiveIdCol] || '').trim();

          if (!hasArchiveId && ARCHIVE_STATUSES.indexOf(status) !== -1) {
            ready++;
          }
        }

        sourceStats.total += total;
        sourceStats.readyToArchive += ready;
        sourceStats.sheets[sheetName] = { total: total, readyToArchive: ready };
      }

      stats.sources[key] = sourceStats;
    } catch (err) {
      stats.sources[key] = { error: err.toString() };
    }
  }

  return {
    success: true,
    stats: stats,
    timestamp: getNow()
  };
}

// ============================================
// ВІДНОВЛЕННЯ ЗАПИСУ З АРХІВУ
// ============================================
// Параметри:
//   archiveSheet: назва архівного аркуша
//   archiveId: ARCHIVE_ID запису
//   user: хто відновлює
// ============================================
function restoreRecord(data) {
  var archiveSheetName = data.archiveSheet;
  var archiveId = data.archiveId;
  var user = data.user || 'system';

  if (!archiveSheetName || !archiveId) {
    return { success: false, error: 'Потрібні archiveSheet та archiveId' };
  }

  // Визначаємо джерело для відновлення
  var sourceKey = ARCHIVE_TO_SOURCE[archiveSheetName];
  if (!sourceKey) {
    return { success: false, error: 'Невідомий архівний аркуш: ' + archiveSheetName };
  }

  var config = SOURCES[sourceKey];
  var meta = ARCHIVE_META[config.type];

  // Відкриваємо архів
  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var archiveSheet = archiveSS.getSheetByName(archiveSheetName);
  if (!archiveSheet) {
    return { success: false, error: 'Архівний аркуш не знайдено' };
  }

  // Шукаємо запис за ARCHIVE_ID
  var lastRow = archiveSheet.getLastRow();
  if (lastRow < 2) {
    return { success: false, error: 'Архів порожній' };
  }

  var archiveColCount = ARCHIVE_TOTAL_COLS[config.type];
  var allValues = archiveSheet.getRange(2, 1, lastRow - 1, archiveColCount).getValues();
  var foundRow = -1;
  var foundData = null;

  for (var i = 0; i < allValues.length; i++) {
    var rowArchiveId = String(allValues[i][meta.ARCHIVE_ID] || '').trim();
    if (rowArchiveId === archiveId) {
      foundRow = i + 2; // 1-based
      foundData = allValues[i];
      break;
    }
  }

  if (foundRow === -1) {
    return { success: false, error: 'Запис не знайдено: ' + archiveId };
  }

  // Визначаємо в який аркуш джерела відновлювати
  var sourceSheetName = String(foundData[meta.SOURCE_SHEET] || '').trim();
  if (!sourceSheetName) {
    // Якщо SOURCE_SHEET не вказано, беремо перший аркуш конфігу
    sourceSheetName = config.sheets[0];
  }

  // Відкриваємо джерело
  var sourceSS;
  try {
    sourceSS = SpreadsheetApp.openById(config.id);
  } catch (err) {
    return { success: false, error: 'Не вдалося відкрити джерело: ' + err.toString() };
  }

  var sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    return { success: false, error: 'Аркуш джерела не знайдено: ' + sourceSheetName };
  }

  // Будуємо рядок для джерела
  var restoredRow = new Array(config.totalCols);
  for (var j = 0; j < config.totalCols; j++) {
    restoredRow[j] = '';
  }

  // Копіюємо дані
  for (var k = 0; k < config.dataCols; k++) {
    restoredRow[k] = foundData[k] !== undefined ? foundData[k] : '';
  }

  // Очищаємо статус архівації — ставимо "work" для посилок або очищаємо
  restoredRow[config.statusCol] = 'work';

  // Додаємо рядок в джерело
  sourceSheet.appendRow(restoredRow);

  // Видаляємо з архіву
  archiveSheet.deleteRow(foundRow);

  // Логуємо
  var recordId = String(foundData[config.idCol] || '');
  writeArchiveLog(
    config.name,
    user,
    'RESTORE',
    recordId,
    sourceSheetName,
    'Відновлено з архіву ' + archiveSheetName,
    'success',
    archiveId
  );

  return {
    success: true,
    restoredTo: {
      spreadsheet: config.name,
      sheet: sourceSheetName
    },
    archiveId: archiveId,
    recordId: recordId,
    timestamp: getNow()
  };
}

// ============================================
// ВИДАЛЕННЯ З АРХІВУ (повне видалення)
// ============================================
// Параметри:
//   archiveSheet: назва архівного аркуша
//   archiveId: ARCHIVE_ID запису (або масив)
//   user: хто видаляє
// ============================================
function deleteFromArchive(data) {
  var archiveSheetName = data.archiveSheet;
  var archiveIds = data.archiveId;
  var user = data.user || 'system';

  if (!archiveSheetName || !archiveIds) {
    return { success: false, error: 'Потрібні archiveSheet та archiveId' };
  }

  // Підтримка одного ID або масиву
  if (typeof archiveIds === 'string') {
    archiveIds = [archiveIds];
  }

  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  var sheet = archiveSS.getSheetByName(archiveSheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + archiveSheetName };
  }

  // Визначаємо тип та мета
  var sourceKey = ARCHIVE_TO_SOURCE[archiveSheetName];
  var config = SOURCES[sourceKey];
  var meta = ARCHIVE_META[config.type];
  var archiveColCount = ARCHIVE_TOTAL_COLS[config.type];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: false, error: 'Архів порожній' };
  }

  var allValues = sheet.getRange(2, 1, lastRow - 1, archiveColCount).getValues();

  // Знаходимо рядки для видалення
  var idSet = {};
  for (var i = 0; i < archiveIds.length; i++) {
    idSet[String(archiveIds[i]).trim()] = true;
  }

  var rowsToDelete = [];
  for (var j = 0; j < allValues.length; j++) {
    var rowArchiveId = String(allValues[j][meta.ARCHIVE_ID] || '').trim();
    if (idSet[rowArchiveId]) {
      rowsToDelete.push(j + 2);
    }
  }

  // Видаляємо знизу вгору
  rowsToDelete.sort(function(a, b) { return b - a; });
  for (var d = 0; d < rowsToDelete.length; d++) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  // Логуємо
  writeArchiveLog(
    'Archive',
    user,
    'DELETE',
    archiveIds.join(', '),
    archiveSheetName,
    'Видалено з архіву: ' + rowsToDelete.length + ' записів',
    'success',
    ''
  );

  return {
    success: true,
    deleted: rowsToDelete.length,
    requested: archiveIds.length,
    timestamp: getNow()
  };
}

// ============================================
// АВТОАРХІВАЦІЯ (для тригеру за часом)
// ============================================
// Встанови тригер: Triggers → Add Trigger → runAutoArchive → Time-driven → Daily
// ============================================
function runAutoArchive() {
  Logger.log('=== АВТОАРХІВАЦІЯ ЗАПУЩЕНА ===');

  // 1. Архівуємо за статусами (все що має статус archived/refused/deleted/transferred)
  var statusResult = archiveByStatus({
    source: 'ALL',
    statuses: ARCHIVE_STATUSES,
    user: 'auto_trigger',
    reason: 'auto_status',
    deleteFromSource: true
  });

  Logger.log('Архівовано за статусом: ' + statusResult.totalArchived);

  // 2. Архівуємо старі записи (старші 60 днів)
  var ageResult = archiveByAge({
    source: 'ALL',
    days: 60,
    user: 'auto_trigger',
    reason: 'auto_age_60d',
    deleteFromSource: true
  });

  Logger.log('Архівовано за віком: ' + ageResult.totalArchived);

  var total = statusResult.totalArchived + ageResult.totalArchived;
  Logger.log('=== ВСЬОГО АРХІВОВАНО: ' + total + ' ===');

  return {
    success: true,
    byStatus: statusResult.totalArchived,
    byAge: ageResult.totalArchived,
    total: total
  };
}

// ============================================
// ЛОГУВАННЯ В АРКУШ "Логи"
// ============================================
// Колонки: A:TIMESTAMP, B:SOURCE, C:USER, D:ACTION,
//          E:RECORD_ID, F:SHEET, G:DETAILS, H:STATUS, I:ARCHIVE_ID
// ============================================
function writeArchiveLog(source, user, action, recordId, sheet, details, status, archiveId) {
  try {
    var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
    var logSheet = archiveSS.getSheetByName(LOG_SHEET_NAME);

    if (!logSheet) {
      logSheet = archiveSS.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow([
        'TIMESTAMP', 'SOURCE', 'USER', 'ACTION',
        'RECORD_ID', 'SHEET', 'DETAILS', 'STATUS', 'ARCHIVE_ID'
      ]);
      logSheet.getRange(1, 1, 1, 9)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    logSheet.appendRow([
      getNow(),
      source || '',
      user || '',
      action || '',
      recordId || '',
      sheet || '',
      details || '',
      status || '',
      archiveId || ''
    ]);
  } catch (e) {
    Logger.log('Log error: ' + e.toString());
  }
}

// ============================================
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Генерація унікального ARCHIVE_ID
function generateArchiveId() {
  var now = new Date();
  var timestamp = Utilities.formatDate(now, 'Europe/Kiev', 'yyyyMMddHHmmss');
  var random = Math.floor(Math.random() * 10000).toString();
  while (random.length < 4) random = '0' + random;
  return 'ARC_' + timestamp + '_' + random;
}

// Поточна дата/час
function getNow() {
  return Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
}

// Перевірка чи рядок порожній
function isEmptyRow(row, colCount) {
  var limit = Math.min(row.length, colCount || row.length);
  for (var c = 0; c < limit; c++) {
    var val = row[c];
    if (val !== '' && val !== null && val !== undefined) {
      return false;
    }
  }
  return true;
}

// Форматування дати
function formatDateValue(value) {
  if (!value) return '';

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }

  return String(value);
}

// Парсинг дати з різних форматів
function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  var str = String(value).trim();
  if (!str) return null;

  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    var d = new Date(str.substring(0, 10));
    return isNaN(d.getTime()) ? null : d;
  }

  // DD.MM.YYYY
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    var parts = str.split('.');
    var d2 = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
    return isNaN(d2.getTime()) ? null : d2;
  }

  // DD/MM/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    var parts2 = str.split('/');
    var d3 = new Date(parseInt(parts2[2]), parseInt(parts2[1]) - 1, parseInt(parts2[0]));
    return isNaN(d3.getTime()) ? null : d3;
  }

  // Спроба стандартного парсингу
  try {
    var d4 = new Date(str);
    return isNaN(d4.getTime()) ? null : d4;
  } catch (e) {
    return null;
  }
}

// Відповідь у форматі JSON
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// МЕНЮ В GOOGLE SHEETS
// ============================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 CRM Архів')
    .addItem('📊 Статистика', 'menuGetStats')
    .addSeparator()
    .addItem('🔄 Архівувати за статусом (всі)', 'menuArchiveByStatus')
    .addItem('📅 Архівувати старі (60+ днів)', 'menuArchiveByAge')
    .addSeparator()
    .addItem('⚡ Автоархівація (повна)', 'runAutoArchive')
    .addToUi();
}

// ============================================
// ТЕСТОВІ / МЕНЮ ФУНКЦІЇ
// ============================================

// Статистика через меню
function menuGetStats() {
  var result = getStats({});

  var msg = '📊 СТАТИСТИКА АРХІВУ\n\n';
  msg += '📁 Всього в архіві: ' + result.stats.total + '\n\n';

  for (var sheet in result.stats.sheets) {
    var s = result.stats.sheets[sheet];
    msg += '  • ' + sheet + ': ' + s.count + ' записів\n';
  }

  msg += '\n📋 Логів: ' + result.stats.logs + '\n\n';
  msg += '═══ ДЖЕРЕЛА ═══\n';

  for (var source in result.stats.sources) {
    var src = result.stats.sources[source];
    if (src.error) {
      msg += '\n❌ ' + source + ': ' + src.error;
    } else {
      msg += '\n📦 ' + source + ':';
      msg += '\n   Всього: ' + src.total + ' | Готові до архіву: ' + src.readyToArchive;
    }
  }

  SpreadsheetApp.getUi().alert(msg);
}

// Архівація за статусом через меню
function menuArchiveByStatus() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '🔄 Архівація за статусом',
    'Архівувати ВСІ записи зі статусами:\n' +
    ARCHIVE_STATUSES.join(', ') + '\n\nз УСІХ джерел?\n\n' +
    'Записи будуть ВИДАЛЕНІ з джерел після архівації.',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  var result = archiveByStatus({
    source: 'ALL',
    statuses: ARCHIVE_STATUSES,
    user: 'menu',
    reason: 'manual_menu',
    deleteFromSource: true
  });

  var msg = '✅ Архівація завершена!\n\nВсього архівовано: ' + result.totalArchived + '\n\n';
  for (var i = 0; i < result.results.length; i++) {
    var r = result.results[i];
    msg += r.name + ': ' + r.archived + ' записів\n';
  }

  ui.alert(msg);
}

// Архівація за віком через меню
function menuArchiveByAge() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '📅 Архівація за віком',
    'Скільки днів? (записи СТАРШІ за вказану кількість днів будуть архівовані)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var days = parseInt(response.getResponseText());
  if (isNaN(days) || days < 1) {
    ui.alert('❌ Введіть коректне число днів');
    return;
  }

  var confirm = ui.alert(
    '⚠️ Підтвердження',
    'Архівувати записи старші ' + days + ' днів з УСІХ джерел?\n\n' +
    'Записи будуть ВИДАЛЕНІ з джерел після архівації.',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  var result = archiveByAge({
    source: 'ALL',
    days: days,
    user: 'menu',
    reason: 'manual_age_' + days + 'd',
    deleteFromSource: true
  });

  ui.alert('✅ Архівовано ' + result.totalArchived + ' записів старших за ' + days + ' днів');
}

// Тест здоров'я
function testHealth() {
  Logger.log('=== CRM ARCHIVE v1.0 ===');
  Logger.log('Archive SS: ' + ARCHIVE_SS_ID);
  Logger.log('Timestamp: ' + getNow());

  // Перевіряємо доступ до всіх таблиць
  for (var key in SOURCES) {
    if (!SOURCES.hasOwnProperty(key)) continue;
    try {
      var ss = SpreadsheetApp.openById(SOURCES[key].id);
      Logger.log('✅ ' + key + ' (' + SOURCES[key].name + '): доступ ОК');
    } catch (e) {
      Logger.log('❌ ' + key + ' (' + SOURCES[key].name + '): ' + e.toString());
    }
  }

  // Перевіряємо архівні аркуші
  var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  for (var sheet in ARCHIVE_SHEETS) {
    if (!ARCHIVE_SHEETS.hasOwnProperty(sheet)) continue;
    var name = ARCHIVE_SHEETS[sheet];
    var s = archiveSS.getSheetByName(name);
    Logger.log(s ? '✅ Архів "' + name + '": OK' : '❌ Архів "' + name + '": НЕ ЗНАЙДЕНО');
  }
}

// Тест статистики
function testStats() {
  var result = getStats({});
  Logger.log('=== СТАТИСТИКА ===');
  Logger.log(JSON.stringify(result, null, 2));
}

// Тест генерації ARCHIVE_ID
function testArchiveId() {
  for (var i = 0; i < 5; i++) {
    Logger.log('ID: ' + generateArchiveId());
    Utilities.sleep(10);
  }
}

// ============================================
// BOTI LOGISTICS — МАРШРУТИ ПОСИЛКИ v2.0
// Apps Script API для таблиці "Маршрути посилки"
// ID: 17g3TFYg11EqdQ9eGrOKQV3n_nqPDFx7dqsJVaGWeDOo
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Маршрути посилки" → Розширення → Apps Script
// 2. Видали весь старий код і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в CRM HTML файл та BOTI Driver
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

var SPREADSHEET_ID = '17g3TFYg11EqdQ9eGrOKQV3n_nqPDFx7dqsJVaGWeDOo';

// Аркуші
var SHEET_LOGS = 'Маршрути водіїв';
var SHEET_MAILING = 'Провірка розсилки';

// Кольори статусів для водіїв
var STATUS_COLORS = {
  'pending':     { bg: '#fffbf0', border: '#ffc107', font: '#ffc107' },
  'in-progress': { bg: '#e3f2fd', border: '#2196F3', font: '#2196F3' },
  'completed':   { bg: '#e8f5e9', border: '#4CAF50', font: '#4CAF50' },
  'cancelled':   { bg: '#ffebee', border: '#dc3545', font: '#dc3545' }
};

// Порядок колонок (A-W = 23 колонки, індекс 0-22)
// Ідентичний до Logistics-Cargo таблиці
// A:Напрямок  B:Номер ТТН  C:Вага  D:Адреса Отримувача
// E:Телефон Отримувача  F:Сума €  G:Статус оплати  H:Оплата
// I:Телефон Реєстратора  J:Примітка  K:Статус посилки  L:ІД
// M:ПІБ  N:дата оформлення  O:Таймінг  P:Примітка смс
// Q:Дата отримання  R:фото  S:Статус  T:Автомобіль
// U:company_id  V:Дата архів  W:ARCHIVE_ID
var COL = {
  DIRECTION: 0,      // A — Напрямок (ua-eu / eu-ua)
  TTN: 1,            // B — Номер ТТН
  WEIGHT: 2,         // C — Вага
  ADDRESS: 3,        // D — Адреса Отримувача
  PHONE: 4,          // E — Телефон Отримувача
  AMOUNT: 5,         // F — Сума €
  PAY_STATUS: 6,     // G — Статус оплати
  PAYMENT: 7,        // H — Оплата
  PHONE_REG: 8,      // I — Телефон Реєстратора
  NOTE: 9,           // J — Примітка
  PARCEL_STATUS: 10, // K — Статус посилки (pending/in-progress/completed/cancelled)
  ID: 11,            // L — ІД
  NAME: 12,          // M — ПІБ
  DATE_REG: 13,      // N — дата оформлення
  TIMING: 14,        // O — Таймінг
  SMS_NOTE: 15,      // P — Примітка смс
  DATE_RECEIVE: 16,  // Q — Дата отримання
  PHOTO: 17,         // R — фото
  STATUS: 18,        // S — Статус (CRM: new/work/archived/refused/deleted)
  VEHICLE: 19,       // T — Автомобіль
  COMPANY_ID: 20,    // U — company_id
  DATE_ARCHIVE: 21,  // V — Дата архів
  ARCHIVE_ID: 22     // W — ARCHIVE_ID
};
var TOTAL_COLS = 23;

// Заголовки для нового аркуша (ідентичні до Logistics-Cargo)
var HEADERS = [
  'Напрямок', 'Номер ТТН', 'Вага', 'Адреса Отримувача',
  'Телефон Отримувача', 'Сума €', 'Статус оплати', 'Оплата',
  'Телефон Реєстратора', 'Примітка', 'Статус посилки', 'ІД',
  'ПІБ', 'дата оформлення', 'Таймінг', 'Примітка смс',
  'Дата отримання', 'фото', 'Статус', 'Автомобіль',
  'company_id', 'Дата архів', 'ARCHIVE_ID'
];

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// ============================================
// Пошук колонки company_id по заголовку (1-й рядок)
// Повертає індекс колонки або -1 якщо не знайдено
// ============================================
function findCompanyIdCol(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === 'company_id') return i;
  }
  return -1;
}

// Маппінг полів API → індексів колонок (ідентичний до Logistics-Cargo)
var FIELD_MAP = {
  direction: COL.DIRECTION,
  ttn: COL.TTN,
  weight: COL.WEIGHT,
  address: COL.ADDRESS,
  phone: COL.PHONE,
  amount: COL.AMOUNT,
  payStatus: COL.PAY_STATUS,
  payment: COL.PAYMENT,
  phoneReg: COL.PHONE_REG,
  note: COL.NOTE,
  parcelStatus: COL.PARCEL_STATUS,
  id: COL.ID,
  name: COL.NAME,
  dateReg: COL.DATE_REG,
  timing: COL.TIMING,
  smsNote: COL.SMS_NOTE,
  dateReceive: COL.DATE_RECEIVE,
  photo: COL.PHOTO,
  status: COL.STATUS,
  vehicle: COL.VEHICLE,
  dateArchive: COL.DATE_ARCHIVE,
  archiveId: COL.ARCHIVE_ID,
  companyId: COL.COMPANY_ID
};

// ============================================
// doGet — ВОДІЇ (BOTI Driver) + Health Check
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'health';
    var sheetParam = (e && e.parameter) ? (e.parameter.sheet || '') : '';
    var companyIdParam = (e && e.parameter) ? (e.parameter.companyId || '') : '';

    switch (action) {
      case 'health':
        return respond({
          success: true,
          version: '1.0',
          service: 'Маршрути Посилки — ЮРА ТРАНСПОРТЕЙШН',
          totalCols: TOTAL_COLS,
          timestamp: new Date().toISOString()
        });

      case 'getDeliveries':
        if (!sheetParam) return respond({ success: false, error: 'Не вказано маршрут (sheet)' });
        return respond(getDeliveries(sheetParam, companyIdParam));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes(companyIdParam));

      default:
        return respond({ success: false, error: 'Невідома GET дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// doPost — CRM + ВОДІЇ
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var payload = data.payload || data;
    // Прокидуємо companyId в payload (фронтенд шле його в data, не в payload)
    payload.companyId = payload.companyId || data.companyId || '';

    switch (action) {
      // --- CRM: МАРШРУТИ ---
      case 'copyToRoute':
        return respond(copyToRoute(payload));

      case 'checkRouteSheets':
        return respond(checkRouteSheets(payload));

      case 'getRoutePassengers':
        return respond(getRoutePackages(payload));

      case 'getRoutePackages':
        return respond(getRoutePackages(payload));

      case 'getAvailableRoutes':
        return respond(getAvailableRoutes(payload.companyId));

      case 'deleteRouteSheet':
        return respond(deleteRouteSheet(payload));

      case 'createRouteSheet':
        return respond(createRouteSheet(payload));

      // --- CRM: ОНОВЛЕННЯ ---
      case 'updateField':
        return respond(updateField(payload));

      case 'updateStatus':
        return respond(updateStatus(payload));

      case 'updateMultiple':
        return respond(updateMultiple(payload));

      // --- CRM: АРХІВАЦІЯ ---
      case 'archivePackages':
        return respond(archiveToExternal(payload));

      case 'restorePackages':
        return respond(changeStatus(payload, 'work'));

      case 'deletePackages':
        return respond(changeStatus(payload, 'deleted'));

      case 'archiveToExternal':
        return respond(archiveToExternal(payload));

      // --- ВОДІЙ: СТАТУС ---
      case 'updateDriverStatus':
      case 'updateStatus_driver':
        return respond(handleDriverStatusUpdate(data));

      // --- РОЗСИЛКА ---
      case 'getMailingStatus':
        return respond(getMailingStatus());

      case 'addMailingRecord':
        return respond(addMailingRecord(payload));

      case 'checkMailing':
        return respond(checkMailingByIds(payload));

      case 'clearMailing':
        return respond(clearMailing(payload));

      case 'clearOldMailing':
        return respond(clearOldMailing(payload));

      // --- СТВОРЕННЯ ---
      case 'addPackage':
        return respond(addPackageToRoute(data));

      // --- ВІДПРАВКА (DISPATCH) ---
      case 'getDispatchItems':
        return respond(getDispatchItems(payload));
      case 'addDispatchItem':
        return respond(addDispatchItem(payload));
      case 'updateDispatchStatus':
        return respond(updateDispatchStatus(data));
      case 'editDispatchItem':
        return respond(editDispatchItem(payload));
      case 'editRoutePackage':
        return respond(editRoutePackage(payload));

      // --- ДЕБАГ ---
      case 'getStructure':
        return respond(getStructure());

      default:
        return respond({ success: false, error: 'Невідома дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// getDeliveries — Читання посилок для водіїв
// ============================================
function getDeliveries(sheetName, companyId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, deliveries: [], count: 0, sheetName: sheetName };
    }

    // Шукаємо колонку company_id по заголовку
    var compIdCol = findCompanyIdCol(sheet);

    var lastCol = sheet.getLastColumn();
    var readCols = Math.max(lastCol, TOTAL_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();
    var deliveries = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Пропускаємо порожні рядки — потрібен хоча б ІД, ТТН або Телефон
      if (!str(row[COL.ID]) && !str(row[COL.TTN]) && !str(row[COL.PHONE]) && !str(row[COL.NAME])) continue;

      // Фільтр по company_id
      if (companyId && compIdCol !== -1) {
        var rowCompanyId = str(row[compIdCol]).toLowerCase();
        if (rowCompanyId !== companyId.toLowerCase()) continue;
      }

      // Пропускаємо архівовані
      var crmStatus = str(row[COL.STATUS]).toLowerCase();
      if (ARCHIVE_STATUSES.indexOf(crmStatus) !== -1) continue;

      deliveries.push({
        rowNum: i + 2,
        ttn: str(row[COL.TTN]),
        weight: str(row[COL.WEIGHT]),
        address: str(row[COL.ADDRESS]),
        direction: str(row[COL.DIRECTION]),
        phone: str(row[COL.PHONE]),
        price: str(row[COL.AMOUNT]),
        paymentStatus: str(row[COL.PAY_STATUS]),
        payment: str(row[COL.PAYMENT]),
        registrarPhone: str(row[COL.PHONE_REG]),
        note: str(row[COL.NOTE]),
        parcelStatus: str(row[COL.PARCEL_STATUS]) || 'pending',
        id: str(row[COL.ID]),
        name: str(row[COL.NAME]),
        createdAt: str(row[COL.DATE_REG]),
        timing: str(row[COL.TIMING]),
        smsNote: str(row[COL.SMS_NOTE]),
        receiveDate: str(row[COL.DATE_RECEIVE]),
        photo: str(row[COL.PHOTO]),
        status: crmStatus || 'new',
        vehicle: str(row[COL.VEHICLE]),
        archiveId: str(row[COL.ARCHIVE_ID]),
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: deliveries.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0 };
    for (var j = 0; j < deliveries.length; j++) {
      var ps = (deliveries[j].parcelStatus || 'pending').toLowerCase();
      if (ps === 'pending') stats.pending++;
      else if (ps === 'in-progress') stats.inProgress++;
      else if (ps === 'completed') stats.completed++;
      else if (ps === 'cancelled') stats.cancelled++;
    }

    return {
      success: true,
      deliveries: deliveries,
      count: deliveries.length,
      sheetName: sheetName,
      stats: stats
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// getRoutePackages — Читання маршруту (для CRM)
// ============================================
function getRoutePackages(payload) {
  try {
    var vehicleName = payload.vehicleName || '';
    var sheetName = payload.sheetName || '';

    if (vehicleName && !sheetName) {
      sheetName = vehicleName;
    }
    if (!sheetName) {
      return { success: false, error: 'Не вказано маршрут' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: true, packages: [], passengers: [], count: 0, sheetName: sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, packages: [], passengers: [], count: 0, sheetName: sheetName, stats: { total: 0 } };
    }

    // Шукаємо колонку company_id по заголовку
    var compIdCol = findCompanyIdCol(sheet);
    var companyId = payload.companyId || '';

    var lastCol = sheet.getLastColumn();
    var readCols = Math.max(lastCol, TOTAL_COLS);
    var dataRange = sheet.getRange(2, 1, lastRow - 1, readCols);
    var data = dataRange.getValues();
    var backgrounds = dataRange.getBackgrounds();
    var packages = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Пропускаємо порожні рядки — потрібен хоча б ІД, ТТН або Телефон
      if (!str(row[COL.ID]) && !str(row[COL.TTN]) && !str(row[COL.PHONE]) && !str(row[COL.NAME])) continue;

      // Фільтр по company_id
      if (companyId && compIdCol !== -1) {
        var rowCompanyId = str(row[compIdCol]).toLowerCase();
        if (rowCompanyId !== companyId.toLowerCase()) continue;
      }

      // Пропускаємо архівовані/видалені/відмовлені
      var crmStatus = str(row[COL.STATUS]).toLowerCase();
      if (ARCHIVE_STATUSES.indexOf(crmStatus) !== -1) continue;

      packages.push({
        rowNum: i + 2,
        ttn: str(row[COL.TTN]),
        weight: str(row[COL.WEIGHT]),
        address: str(row[COL.ADDRESS]),
        direction: str(row[COL.DIRECTION]),
        phone: str(row[COL.PHONE]),
        amount: str(row[COL.AMOUNT]),
        payStatus: str(row[COL.PAY_STATUS]),
        payment: str(row[COL.PAYMENT]),
        phoneReg: str(row[COL.PHONE_REG]),
        note: str(row[COL.NOTE]),
        parcelStatus: str(row[COL.PARCEL_STATUS]) || 'pending',
        id: str(row[COL.ID]),
        name: str(row[COL.NAME]),
        dateReg: str(row[COL.DATE_REG]),
        timing: str(row[COL.TIMING]),
        smsNote: str(row[COL.SMS_NOTE]),
        dateReceive: str(row[COL.DATE_RECEIVE]),
        photo: str(row[COL.PHOTO]),
        status: str(row[COL.STATUS]),
        vehicle: str(row[COL.VEHICLE]),
        archiveId: str(row[COL.ARCHIVE_ID]),
        rowColor: backgrounds[i][0],
        sheet: sheetName
      });
    }

    // Статистика
    var stats = { total: packages.length, pending: 0, inProgress: 0, completed: 0, cancelled: 0 };
    for (var j = 0; j < packages.length; j++) {
      var ps = (packages[j].parcelStatus || 'pending').toLowerCase();
      if (ps === 'pending') stats.pending++;
      else if (ps === 'in-progress') stats.inProgress++;
      else if (ps === 'completed') stats.completed++;
      else if (ps === 'cancelled') stats.cancelled++;
    }

    return {
      success: true,
      packages: packages,
      passengers: packages,  // alias для CRM сумісності
      count: packages.length,
      sheetName: sheetName,
      vehicleName: vehicleName,
      stats: stats
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// getAvailableRoutes — Список маршрутних аркушів
// ============================================
function getAvailableRoutes(companyId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ss.getSheets();
    var routes = [];

    // Службові аркуші які НЕ є маршрутами
    var excludePatterns = ['логи', 'logs', 'водіїв', 'розсилк', 'провірка', 'перевірка', 'template', 'шаблон', 'тест', 'test', 'архів', 'маршрути водіїв', 'відпр'];

    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      var nameLower = name.toLowerCase();

      // Пропускаємо службові аркуші
      var isExcluded = false;
      for (var e = 0; e < excludePatterns.length; e++) {
        if (nameLower.indexOf(excludePatterns[e]) !== -1) {
          isExcluded = true;
          break;
        }
      }
      if (isExcluded) continue;

      var lastRow = sheets[i].getLastRow();

      // Порожній аркуш (тільки заголовок або зовсім порожній) — показуємо з count=0
      if (lastRow < 2) {
        routes.push({
          name: name,
          vehicle: name,
          count: 0,
          sheetId: sheets[i].getSheetId()
        });
        continue;
      }

      // Фільтр по company_id — показуємо тільки маршрути де є рядки з цим company_id
      // Виключаємо архівовані записи з підрахунку
      var statusCol = COL.STATUS; // Колонка статусу
      var allData = sheets[i].getRange(2, 1, lastRow - 1, Math.max(statusCol + 1, (companyId ? sheets[i].getLastColumn() : statusCol + 1))).getValues();
      var compIdCol = companyId ? findCompanyIdCol(sheets[i]) : -1;
      if (companyId && compIdCol === -1) continue; // Немає колонки company_id — не показуємо

      var count = 0;
      for (var r = 0; r < allData.length; r++) {
        var rowStatus = String(allData[r][statusCol] || '').trim().toLowerCase();
        if (ARCHIVE_STATUSES.indexOf(rowStatus) !== -1) continue; // Пропускаємо архівовані
        if (companyId) {
          if (String(allData[r][compIdCol] || '').trim().toLowerCase() !== companyId.toLowerCase()) continue;
        }
        count++;
      }
      // Маршрут показуємо навіть якщо count=0 (порожній або все заархівовано) — щоб можна було додати нові посилки

      routes.push({
        name: name,
        vehicle: name,
        count: count,
        sheetId: sheets[i].getSheetId()
      });
    }

    return { success: true, routes: routes, count: routes.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// copyToRoute — CRM → маршрутний аркуш
// ============================================
function copyToRoute(payload) {
  try {
    var packagesByVehicle = payload.packagesByVehicle;
    var conflictAction = payload.conflictAction || 'add';

    if (!packagesByVehicle) {
      return { success: false, error: 'Немає посилок' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var totalCopied = 0;
    var totalArchived = 0;
    var totalCleared = 0;
    var results = [];

    for (var vehicleName in packagesByVehicle) {
      if (!packagesByVehicle.hasOwnProperty(vehicleName)) continue;
      var packages = packagesByVehicle[vehicleName];
      if (!packages || !packages.length) continue;

      var sheetName = vehicleName;
      var sheet = ss.getSheetByName(sheetName);

      // Створюємо аркуш якщо не існує
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setBackground('#1a1a2e')
          .setFontColor('#ffffff')
          .setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      // Обробка конфліктів
      var lastRow = sheet.getLastRow();
      if (lastRow > 1 && conflictAction !== 'add') {
        if (conflictAction === 'clear') {
          totalCleared += lastRow - 1;
          sheet.deleteRows(2, lastRow - 1);
        } else if (conflictAction === 'archive') {
          totalArchived += lastRow - 1;
          // Архівуємо старі рядки (ставимо статус)
          var archiveRange = sheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1);
          var archiveValues = [];
          for (var a = 0; a < lastRow - 1; a++) archiveValues.push(['archived']);
          archiveRange.setValues(archiveValues);

          var dateRange = sheet.getRange(2, COL.DATE_ARCHIVE + 1, lastRow - 1, 1);
          var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
          var dateValues = [];
          for (var d = 0; d < lastRow - 1; d++) dateValues.push([dateNow]);
          dateRange.setValues(dateValues);
        }
      }

      // Записуємо посилки
      var rows = [];
      for (var p = 0; p < packages.length; p++) {
        var pkg = packages[p];
        var newRow = new Array(TOTAL_COLS);
        for (var c = 0; c < TOTAL_COLS; c++) newRow[c] = '';

        newRow[COL.DIRECTION] = pkg.direction || '';
        newRow[COL.TTN] = pkg.ttn || '';
        newRow[COL.WEIGHT] = pkg.weight || '';
        newRow[COL.ADDRESS] = pkg.address || '';
        newRow[COL.PHONE] = pkg.phone || '';
        newRow[COL.AMOUNT] = pkg.amount || '';
        newRow[COL.PAY_STATUS] = pkg.payStatus || '';
        newRow[COL.PAYMENT] = pkg.payment || '';
        newRow[COL.PHONE_REG] = pkg.phoneReg || '';
        newRow[COL.NOTE] = pkg.note || '';
        newRow[COL.PARCEL_STATUS] = 'pending';
        newRow[COL.ID] = pkg.id || '';
        newRow[COL.NAME] = pkg.name || '';
        newRow[COL.DATE_REG] = pkg.dateReg || '';
        newRow[COL.TIMING] = pkg.timing || '';
        newRow[COL.SMS_NOTE] = pkg.smsNote || '';
        newRow[COL.DATE_RECEIVE] = pkg.dateReceive || '';
        newRow[COL.PHOTO] = pkg.photo || '';
        newRow[COL.STATUS] = 'new';
        newRow[COL.VEHICLE] = vehicleName;
        newRow[COL.COMPANY_ID] = payload.companyId || '';

        rows.push(newRow);
      }

      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setValues(rows);

        // Фарбуємо pending рядки
        var pendingColors = STATUS_COLORS['pending'];
        if (pendingColors) {
          sheet.getRange(startRow, 1, rows.length, TOTAL_COLS).setBackground(pendingColors.bg);
        }

        totalCopied += rows.length;
      }

      results.push({ vehicle: vehicleName, sheet: sheetName, copied: packages.length });

      // Авторозмір
      try { sheet.autoResizeColumns(1, Math.min(20, TOTAL_COLS)); } catch (e) {}
    }

    writeLog('copyToRoute', 'bulk', 0, 'copied: ' + totalCopied,
      'archived: ' + totalArchived + ' cleared: ' + totalCleared);

    return {
      success: true,
      copied: totalCopied,
      archived: totalArchived,
      cleared: totalCleared,
      details: results
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// checkRouteSheets — Перевірка чи є дані
// ============================================
function checkRouteSheets(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var vehicleNames = payload.vehicleNames || [];
    var existing = [];

    for (var i = 0; i < vehicleNames.length; i++) {
      var sheetName = vehicleNames[i];
      var sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        existing.push({
          vehicle: vehicleNames[i],
          sheet: sheetName,
          count: sheet.getLastRow() - 1
        });
      }
    }

    return { success: true, existing: existing };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// createRouteSheet — Створити маршрутний аркуш
// ============================================
function createRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = vehicleName;

    var existing = ss.getSheetByName(sheetName);
    if (existing) {
      return { success: true, sheetName: sheetName, existed: true };
    }

    var sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    writeLog('createRouteSheet', sheetName, 0, 'created', '');

    return { success: true, sheetName: sheetName, vehicleName: vehicleName };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// deleteRouteSheet — Видалити маршрутний аркуш
// ============================================
function deleteRouteSheet(payload) {
  try {
    var vehicleName = payload.vehicleName;
    if (!vehicleName) return { success: false, error: 'Не вказано назву авто' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = vehicleName;
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return { success: true, message: 'Аркуш не існує', deleted: false };
    }

    var rowCount = sheet.getLastRow() - 1;
    if (rowCount > 0 && !payload.force) {
      return {
        success: false,
        error: 'Аркуш містить ' + rowCount + ' записів. Використайте force=true',
        recordsCount: rowCount
      };
    }

    ss.deleteSheet(sheet);
    writeLog('deleteRouteSheet', sheetName, 0, 'deleted', '');

    return { success: true, sheetName: sheetName, deleted: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// updateField — Оновити одне поле
// ============================================
function updateField(payload) {
  var sheetName = payload.sheet;
  var rowNum = parseInt(payload.rowNum);
  var field = payload.field;
  var value = payload.value;

  if (!sheetName || !rowNum || !field) {
    return { success: false, error: 'Відсутні sheet, rowNum або field' };
  }
  if (!FIELD_MAP.hasOwnProperty(field)) {
    return { success: false, error: 'Невідоме поле: ' + field };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  // Верифікація
  if (payload.expectedId) {
    var currentId = str(sheet.getRange(rowNum, COL.ID + 1).getValue());
    if (currentId !== String(payload.expectedId).trim()) {
      return {
        success: false,
        error: 'conflict',
        message: 'Рядок змінився. Очікувався ІД: ' + payload.expectedId + ', фактичний: ' + currentId
      };
    }
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);
  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updateStatus — Змінити CRM статус
// ============================================
function updateStatus(payload) {
  var sheetName = payload.sheet;
  var rowNum = parseInt(payload.rowNum);
  var newStatus = payload.status;

  if (!sheetName || !rowNum || !newStatus) {
    return { success: false, error: 'Відсутні sheet, rowNum або status' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, error: 'Аркуш не знайдено' };
  if (rowNum > sheet.getLastRow()) return { success: false, error: 'Рядок не існує' };

  var oldStatus = str(sheet.getRange(rowNum, COL.STATUS + 1).getValue());
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // При архівації — записуємо дату
  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, oldStatus + ' → ' + newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus };
}

// ============================================
// updateMultiple — Масове оновлення полів
// ============================================
function updateMultiple(payload) {
  var items = payload.items || payload.packages || [];
  var updated = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (!item.sheet || !item.rowNum) {
      errors.push('Елемент ' + i + ': відсутні sheet/rowNum');
      continue;
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(item.sheet);
    if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

    var rowNum = parseInt(item.rowNum);
    if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum + ': не існує'); continue; }

    for (var field in item) {
      if (item.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
        if (item[field] !== undefined) {
          sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(item[field]);
        }
      }
    }
    updated++;
  }

  return { success: true, updated: updated, errors: errors.length > 0 ? errors : undefined };
}

// ============================================
// changeStatus — Масова зміна CRM статусу
// ============================================
function changeStatus(payload, newStatus) {
  try {
    var items = payload.packages || payload.items || payload.passengers || [];
    var note = payload.note || '';

    if (items.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var user = payload.user || 'crm';
    var changed = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheet = ss.getSheetByName(item.sheet);
      if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (!rowNum || rowNum < 2 || rowNum > sheet.getLastRow()) {
        errors.push('Рядок ' + rowNum + ': не існує');
        continue;
      }

      sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

      if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
        var companyId = payload.companyId || '';
        sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
        // Записуємо companyId щоб точно був в рядку
        if (companyId) {
          sheet.getRange(rowNum, COL.COMPANY_ID + 1).setValue(companyId);
        }
      }

      if (note) {
        var currentNote = str(sheet.getRange(rowNum, COL.NOTE + 1).getValue());
        var updatedNote = note + (currentNote ? ' | ' + currentNote : '');
        sheet.getRange(rowNum, COL.NOTE + 1).setValue(updatedNote);
      }

      changed++;
    }

    writeLog('changeStatus:' + newStatus, 'bulk', 0, changed + '/' + items.length, note || '');

    return {
      success: true,
      changed: changed,
      total: items.length,
      status: newStatus,
      errors: errors.length > 0 ? errors : undefined
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// archiveToExternal — Архівація маршрутних посилок
// In-place: оновлює STATUS, DATE_ARCHIVE, ARCHIVE_ID
// без копіювання в зовнішню таблицю
// ============================================
function archiveToExternal(payload) {
  try {
    var items = payload.items || payload.packages || [];
    var companyId = payload.companyId || '';
    var user = companyId ? ('archived_' + companyId) : (payload.user || 'crm');
    var reason = payload.reason || 'route_archive';

    if (items.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var count = 0;
    var errors = [];

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var sheet = ss.getSheetByName(item.sheet);
      if (!sheet) { errors.push(item.sheet + ': не знайдено'); continue; }

      var rowNum = parseInt(item.rowNum);
      if (rowNum > sheet.getLastRow()) { errors.push('Рядок ' + rowNum); continue; }

      var rowData = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];
      var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
      if (existingArchiveId) {
        errors.push('Рядок ' + rowNum + ': вже архівовано');
        continue;
      }

      var archiveId = generateArchiveId_();

      // Визначаємо правильний статус: refused/deleted/archived залежно від reason
      var archiveStatus = 'archived';
      if (reason === 'refused' || reason === 'deleted' || reason === 'transferred') {
        archiveStatus = reason;
      }

      // In-place: оновлюємо 3 колонки в тій самій таблиці
      sheet.getRange(rowNum, COL.STATUS + 1).setValue(archiveStatus);
      sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
      sheet.getRange(rowNum, COL.ARCHIVE_ID + 1).setValue(archiveId);
      count++;
    }

    writeLog('archivePackages', 'bulk', 0, 'archived: ' + count,
      count + '/' + items.length + ' архівовано in-place | by: ' + user + ' | reason: ' + reason);

    return {
      success: true,
      count: count,
      total: items.length,
      errors: errors.length > 0 ? errors : undefined
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Генерація ARCHIVE_ID
function generateArchiveId_() {
  var now = new Date();
  var ts = Utilities.formatDate(now, 'Europe/Kiev', 'yyyyMMddHHmmss');
  var rnd = Math.floor(Math.random() * 10000).toString();
  while (rnd.length < 4) rnd = '0' + rnd;
  return 'ARC_' + ts + '_' + rnd;
}

// ============================================
// handleDriverStatusUpdate — Оновлення від водія
// ============================================
function handleDriverStatusUpdate(data) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Логуємо в аркуш водіїв
    var logSheet = ss.getSheetByName(SHEET_LOGS);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOGS);
      logSheet.getRange(1, 1, 1, 10).setValues([[
        'Дата', 'Час', 'Водій', 'Маршрут', 'Номер посилки',
        'Адреса', 'Статус', 'Причина скасування', 'Телефон', 'Сума'
      ]]);
      logSheet.getRange(1, 1, 1, 10)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }

    var now = new Date();
    logSheet.appendRow([
      Utilities.formatDate(now, 'Europe/Kiev', 'yyyy-MM-dd'),
      Utilities.formatDate(now, 'Europe/Kiev', 'HH:mm:ss'),
      data.driverId || '',
      data.routeName || '',
      data.deliveryNumber || '',
      data.address || '',
      data.status || '',
      data.cancelReason || '',
      data.phone || '',
      data.price || ''
    ]);

    // 2. Оновлюємо статус у маршрутному аркуші
    var routeSheet = ss.getSheetByName(data.routeName);
    if (!routeSheet) {
      return { success: true, message: 'Логовано (маршрут не знайдено)' };
    }

    var allData = routeSheet.getDataRange().getValues();
    var rowsUpdated = 0;
    var deliveryNumber = str(data.deliveryNumber);
    var deliveryId = str(data.deliveryId || data.id || '');

    for (var i = 1; i < allData.length; i++) {
      // Шукаємо по ІД, ТТН або номеру телефону
      var rowId = str(allData[i][COL.ID]);
      var rowTtn = str(allData[i][COL.TTN]);
      var matched = (deliveryId && rowId === deliveryId) ||
                    (deliveryNumber && rowTtn === deliveryNumber) ||
                    (deliveryNumber && rowId === deliveryNumber);
      if (matched) {
        var rowNum = i + 1;

        // Записуємо статус посилки (M — parcelStatus)
        routeSheet.getRange(rowNum, COL.PARCEL_STATUS + 1).setValue(data.status);

        // Якщо completed — записуємо дату отримання
        if (data.status === 'completed') {
          routeSheet.getRange(rowNum, COL.DATE_RECEIVE + 1).setValue(
            Utilities.formatDate(now, 'Europe/Kiev', 'yyyy-MM-dd HH:mm')
          );
        }

        // Причина скасування → в Примітку
        if (data.status === 'cancelled' && data.cancelReason) {
          var currentNote = str(routeSheet.getRange(rowNum, COL.NOTE + 1).getValue());
          var newNote = 'Скасовано: ' + data.cancelReason + (currentNote ? ' | ' + currentNote : '');
          routeSheet.getRange(rowNum, COL.NOTE + 1).setValue(newNote);
        }

        // Кольори
        var colors = STATUS_COLORS[data.status];
        if (colors) {
          var readCols = Math.min(routeSheet.getLastColumn(), TOTAL_COLS);
          var rangeToColor = routeSheet.getRange(rowNum, 1, 1, readCols);
          rangeToColor.setBackground(colors.bg);
          rangeToColor.setBorder(true, true, true, true, true, true,
            colors.border, SpreadsheetApp.BorderStyle.SOLID);

          var statusCell = routeSheet.getRange(rowNum, COL.PARCEL_STATUS + 1);
          statusCell.setFontColor(colors.font);
          statusCell.setFontWeight('bold');
        }

        rowsUpdated++;
      }
    }

    if (rowsUpdated === 0) {
      return { success: true, message: 'Логовано (посилку не знайдено в маршруті)' };
    }

    return {
      success: true,
      message: 'Статус записано',
      updatedRows: rowsUpdated
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================
// РОЗСИЛКА — "Провірка розсилки"
//
// Структура аркуша (9 колонок A-I):
//   A: Стовпець 1 (ІД посилки)
//   B: Стовпець 2 (користувач/дата)
//   C: Стовпець 3 (дод. інфо)
//   D: Статус (sent / archived)
//   E: DATE_ARCHIVE
//   F: ARCHIVED_BY
//   G: ARCHIVE_REASON
//   H: SOURCE_SHEET
//   I: ARCHIVE_ID
//
// Логіка:
// - SmartSender кидає ІД в цей аркуш
// - checkMailing → перевіряє чи ІД є (розсилка зроблена)
// - Записи з ARCHIVE_ID (кол. I) = вже архівовані → НЕ актуальні
// - clearOldMailing → чистить записи старіші за N днів
// - clearMailing → повне очищення (або тільки архівованих)
// ============================================

var MAILING_COL = {
  ID: 0,              // A — ІД посилки
  USER: 1,            // B — хто/коли
  INFO: 2,            // C — дод. інфо
  STATUS: 3,          // D — Статус
  DATE_ARCHIVE: 4,    // E — DATE_ARCHIVE
  ARCHIVED_BY: 5,     // F — ARCHIVED_BY
  ARCHIVE_REASON: 6,  // G — ARCHIVE_REASON
  SOURCE_SHEET: 7,    // H — SOURCE_SHEET
  ARCHIVE_ID: 8       // I — ARCHIVE_ID
};
var MAILING_COLS = 9;

// getMailingStatus — Всі ІД з розсилки
// Повертає activeIds (без ARCHIVE_ID) + allIds (всі)
function getMailingStatus() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, mailingIds: [], activeIds: [], archivedIds: [], count: 0 };
    }

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    var activeIds = [];
    var archivedIds = [];
    var allData = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var id = str(row[MAILING_COL.ID]);
      if (!id || id.indexOf('dd.mm') !== -1) continue;

      var archiveId = readCols > MAILING_COL.ARCHIVE_ID ? str(row[MAILING_COL.ARCHIVE_ID]) : '';
      var status = readCols > MAILING_COL.STATUS ? str(row[MAILING_COL.STATUS]) : '';
      var isArchived = archiveId !== '' || status === 'archived';

      allData.push({
        rowNum: i + 2,
        id: id,
        user: str(row[MAILING_COL.USER]),
        status: status,
        archiveId: archiveId,
        isArchived: isArchived
      });

      if (isArchived) {
        archivedIds.push(id);
      } else {
        activeIds.push(id);
      }
    }

    // Унікальні активні
    var uniqueActive = [];
    var seen = {};
    for (var u = 0; u < activeIds.length; u++) {
      if (!seen[activeIds[u]]) {
        seen[activeIds[u]] = true;
        uniqueActive.push(activeIds[u]);
      }
    }

    // Унікальні всі (для сумісності)
    var allIds = [];
    var seenAll = {};
    for (var a = 0; a < allData.length; a++) {
      if (!seenAll[allData[a].id]) {
        seenAll[allData[a].id] = true;
        allIds.push(allData[a].id);
      }
    }

    return {
      success: true,
      mailingIds: allIds,          // всі (для сумісності)
      activeIds: uniqueActive,     // тільки актуальні (без ARCHIVE_ID)
      archivedIds: archivedIds,
      mailingData: allData,
      count: allIds.length,
      activeCount: uniqueActive.length,
      archivedCount: archivedIds.length
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// checkMailingByIds — Перевірити конкретні ІД
// Дивиться тільки АКТУАЛЬНІ записи (без ARCHIVE_ID)
function checkMailingByIds(payload) {
  try {
    var idsToCheck = payload.ids || [];
    if (idsToCheck.length === 0) {
      return { success: true, results: {}, count: 0 };
    }

    var statusResult = getMailingStatus();
    if (!statusResult.success) return statusResult;

    // Використовуємо activeIds — тільки актуальні
    var activeSet = {};
    for (var m = 0; m < statusResult.activeIds.length; m++) {
      activeSet[statusResult.activeIds[m]] = true;
    }

    // Також маємо allIds для інфо
    var allSet = {};
    for (var a = 0; a < statusResult.mailingIds.length; a++) {
      allSet[statusResult.mailingIds[a]] = true;
    }

    var results = {};
    var mailedCount = 0;
    var archivedMailedCount = 0;

    for (var i = 0; i < idsToCheck.length; i++) {
      var id = String(idsToCheck[i]).trim();
      var isActive = activeSet[id] === true;
      var wasEverMailed = allSet[id] === true;

      results[id] = {
        mailed: isActive,                // актуальна розсилка
        wasMailedBefore: wasEverMailed,   // колись була (навіть якщо архівована)
        archivedMailing: wasEverMailed && !isActive  // розсилка була, але запис архівований
      };

      if (isActive) mailedCount++;
      if (wasEverMailed && !isActive) archivedMailedCount++;
    }

    return {
      success: true,
      results: results,
      total: idsToCheck.length,
      mailed: mailedCount,
      archivedMailed: archivedMailedCount,
      notMailed: idsToCheck.length - mailedCount - archivedMailedCount
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// addMailingRecord — Записати що розсилка зроблена
function addMailingRecord(payload) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_MAILING);
      sheet.getRange(1, 1, 1, MAILING_COLS).setValues([[
        'Стовпець 1', 'Стовпець 2', 'Стовпець 3', 'Статус',
        'DATE_ARCHIVE', 'ARCHIVED_BY', 'ARCHIVE_REASON', 'SOURCE_SHEET', 'ARCHIVE_ID'
      ]]);
    }

    var records = payload.records || [];
    var userName = payload.userName || 'CRM';

    if (records.length === 0) {
      return { success: false, error: 'Немає записів' };
    }

    var now = new Date();
    var date = Utilities.formatDate(now, 'Europe/Kiev', 'dd.MM.yyyy');
    var rows = [];

    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var newRow = new Array(MAILING_COLS);
      for (var c = 0; c < MAILING_COLS; c++) newRow[c] = '';
      newRow[MAILING_COL.ID] = r.id || r.packageId || '';
      newRow[MAILING_COL.USER] = userName;
      newRow[MAILING_COL.INFO] = date;
      newRow[MAILING_COL.STATUS] = r.status || 'sent';
      rows.push(newRow);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, MAILING_COLS).setValues(rows);

    return { success: true, added: rows.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// clearMailing — Очищення розсилки
// mode: 'all' — видалити все, 'archived' — тільки з ARCHIVE_ID
function clearMailing(payload) {
  try {
    var mode = payload.mode || 'archived';  // за замовчуванням — тільки архівовані
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, deleted: 0, message: 'Нічого очищувати' };
    }

    if (mode === 'all') {
      // Видаляємо ВСЕ крім заголовка
      var totalRows = sheet.getLastRow() - 1;
      sheet.deleteRows(2, totalRows);
      writeLog('clearMailing', SHEET_MAILING, 0, 'all', totalRows + ' рядків');
      return { success: true, deleted: totalRows, mode: 'all' };
    }

    // mode === 'archived' — видаляємо тільки рядки з ARCHIVE_ID
    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    // Збираємо рядки для видалення (знизу вгору)
    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var archiveId = readCols > MAILING_COL.ARCHIVE_ID ? str(data[i][MAILING_COL.ARCHIVE_ID]) : '';
      var status = readCols > MAILING_COL.STATUS ? str(data[i][MAILING_COL.STATUS]) : '';
      if (archiveId || status === 'archived') {
        rowsToDelete.push(i + 2);
      }
    }

    // Видаляємо знизу вгору
    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    writeLog('clearMailing', SHEET_MAILING, 0, 'archived', rowsToDelete.length + ' рядків');

    return {
      success: true,
      deleted: rowsToDelete.length,
      remaining: (lastRow - 1) - rowsToDelete.length,
      mode: 'archived'
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// clearOldMailing — Очистити записи старіші за N днів
function clearOldMailing(payload) {
  try {
    var days = parseInt(payload.days) || 30;  // за замовчуванням 30 днів
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_MAILING);

    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, deleted: 0, message: 'Нічого очищувати' };
    }

    var lastRow = sheet.getLastRow();
    var readCols = Math.min(sheet.getLastColumn(), MAILING_COLS);
    var data = sheet.getRange(2, 1, lastRow - 1, readCols).getValues();

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - days);

    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 0; i--) {
      // Дата може бути в колонці C (INFO) або E (DATE_ARCHIVE)
      var dateStr = str(data[i][MAILING_COL.INFO]);
      var archiveDateStr = readCols > MAILING_COL.DATE_ARCHIVE ? str(data[i][MAILING_COL.DATE_ARCHIVE]) : '';

      var recordDate = parseMailingDate(dateStr) || parseMailingDate(archiveDateStr);

      if (recordDate && recordDate < cutoff) {
        rowsToDelete.push(i + 2);
      }
    }

    // Видаляємо знизу вгору
    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    writeLog('clearOldMailing', SHEET_MAILING, 0, '>' + days + ' днів', rowsToDelete.length + ' рядків');

    return {
      success: true,
      deleted: rowsToDelete.length,
      remaining: (lastRow - 1) - rowsToDelete.length,
      days: days
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Парсинг дати розсилки (dd.MM.yyyy або yyyy-MM-dd)
function parseMailingDate(dateStr) {
  if (!dateStr) return null;
  // dd.MM.yyyy
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(dateStr)) {
    var p = dateStr.split('.');
    return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
  }
  // yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) {
    return new Date(dateStr.substring(0, 10));
  }
  try {
    var d = new Date(dateStr);
    if (!isNaN(d.getTime())) return d;
  } catch (e) {}
  return null;
}

// ============================================
// getStructure — Дебаг
// ============================================
function getStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var result = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

    result.push({
      sheet: sheet.getName(),
      rows: lastRow,
      cols: lastCol,
      headers: headers
    });
  }

  return { success: true, sheets: result };
}


// ============================================
// addPackageToRoute — Додати посилку з Drivers UI
// ============================================
function addPackageToRoute(data) {
  try {
    var fields = data.fields || {};
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = data.sheet || fields.vehicle;
    if (!sheetName) {
      return { success: false, error: 'Не вказано маршрут (sheet/vehicle)' };
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var today = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
    var newId = 'drv_' + new Date().getTime();

    var newRow = new Array(TOTAL_COLS);
    for (var i = 0; i < TOTAL_COLS; i++) newRow[i] = '';

    // Маппінг полів через FIELD_MAP
    for (var field in fields) {
      if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
        newRow[FIELD_MAP[field]] = fields[field];
      }
    }

    if (!newRow[COL.DATE_REG]) newRow[COL.DATE_REG] = today;
    if (!newRow[COL.ID]) newRow[COL.ID] = newId;
    if (!newRow[COL.STATUS]) newRow[COL.STATUS] = 'new';

    sheet.appendRow(newRow);
    var newRowNum = sheet.getLastRow();

    // Записуємо company_id якщо колонка є
    var companyId = data.companyId || '';
    if (companyId) {
      var compCol = findCompanyIdCol(sheet);
      if (compCol >= 0) {
        sheet.getRange(newRowNum, compCol + 1).setValue(companyId);
      }
    }

    writeLog('addPackage', sheetName, newRowNum, 'new',
      'ПіБ: ' + (fields.name || '') + ' | ТТН: ' + (fields.ttn || '') + ' | Тел: ' + (fields.phone || '') + ' | Driver UI');

    return {
      success: true,
      sheet: sheetName,
      rowNum: newRowNum,
      id: newId
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================
// ЛОГУВАННЯ — пише в архівну таблицю, аркуш "Логи"
// ============================================
var ARCHIVE_SS_ID_LOG = '1Kmf6NF1sJUi-j3SamrhUqz337pcZSvZCUkGxBzari6U';

function writeLog(action, sheetName, rowNum, detail, extra) {
  try {
    var archiveSS = SpreadsheetApp.openById(ARCHIVE_SS_ID_LOG);
    var logSheet = archiveSS.getSheetByName('Логи');

    if (!logSheet) {
      logSheet = archiveSS.insertSheet('Логи');
      logSheet.appendRow(['Дата/Час', 'Дія', 'Аркуш', 'Рядок', 'Деталі', 'Дані']);
      logSheet.getRange(1, 1, 1, 6)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([timestamp, action, sheetName, rowNum, detail, extra || '']);
  } catch (e) {
    Logger.log('Log error: ' + e.toString());
  }
}

// ============================================
// ДОПОМІЖНІ
// ============================================

// Безпечне перетворення в string
function str(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }
  return String(value).trim();
}

// JSON відповідь
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
  ui.createMenu('Маршрути Посилки')
    .addItem('Список маршрутів', 'menuAvailableRoutes')
    .addItem('Структура таблиці', 'menuStructure')
    .addToUi();
}

function menuAvailableRoutes() {
  var result = getAvailableRoutes();
  var msg = 'Маршрутів: ' + result.count + '\n\n';
  for (var i = 0; i < result.routes.length; i++) {
    var r = result.routes[i];
    msg += r.name + ' — ' + r.count + ' записів\n';
  }
  SpreadsheetApp.getUi().alert('Маршрути', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuStructure() {
  var result = getStructure();
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('[' + s.sheet + '] ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('  Headers: ' + s.headers.join(' | '));
  }
  SpreadsheetApp.getUi().alert('Дивись Logger (Ctrl+Enter)');
}


// ============================================
// ВІДПРАВКА (DISPATCH) — аркуші "(відпр)"
// Колонки: A:ПІБ  B:№ посилки  C:Місто/Нова Пошта
//          D:Опис/фото  E:Вага  F:Сума  G:Тип оплати
//          H:Валюта  I:Конверт  J:ID/Телефон відправника
//          K:company_id
// ============================================

var DCOL = {
  NAME: 0,          // A — ПІБ
  PARCEL_NUM: 1,    // B — № посилки
  CITY: 2,          // C — Місто/Нова Пошта
  DESCRIPTION: 3,   // D — Опис/фото
  WEIGHT: 4,        // E — Вага
  AMOUNT: 5,        // F — Сума
  PAYMENT_TYPE: 6,  // G — Тип оплати
  CURRENCY: 7,      // H — Валюта
  ENVELOPE: 8,      // I — Конверт
  SENDER_PHONE: 9,  // J — ID/Телефон відправника
  COMPANY_ID: 10    // K — company_id
};
var DISPATCH_TOTAL_COLS = 11;

// ============================================
// getDispatchItems — Читання відправок з аркуша "(відпр)"
// ============================================
function getDispatchItems(payload) {
  try {
    var sheetName = payload.sheetName || '';
    if (!sheetName) {
      return { success: false, error: 'Не вказано аркуш відправки' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, items: [], count: 0, sheetName: sheetName };
    }

    var lastCol = sheet.getLastColumn();
    var numCols = Math.max(lastCol, DISPATCH_TOTAL_COLS);
    var range = sheet.getRange(2, 1, lastRow - 1, numCols);
    var data = range.getValues();

    var companyId = payload.companyId || '';
    var items = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // Пропускаємо порожні рядки
      if (!row[DCOL.NAME] && !row[DCOL.PARCEL_NUM] && !row[DCOL.SENDER_PHONE]) continue;

      // Фільтр по company_id
      if (companyId && row[DCOL.COMPANY_ID] && String(row[DCOL.COMPANY_ID]).trim() !== companyId) continue;

      items.push({
        rowNum: i + 2,
        name: String(row[DCOL.NAME] || '').trim(),
        parcelNum: String(row[DCOL.PARCEL_NUM] || '').trim(),
        city: String(row[DCOL.CITY] || '').trim(),
        description: String(row[DCOL.DESCRIPTION] || '').trim(),
        weight: String(row[DCOL.WEIGHT] || '').trim(),
        amount: String(row[DCOL.AMOUNT] || '').trim(),
        paymentType: String(row[DCOL.PAYMENT_TYPE] || '').trim(),
        currency: String(row[DCOL.CURRENCY] || '').trim(),
        envelope: String(row[DCOL.ENVELOPE] || '').trim(),
        senderPhone: String(row[DCOL.SENDER_PHONE] || '').trim(),
        driverStatus: 'pending'
      });
    }

    return {
      success: true,
      items: items,
      count: items.length,
      sheetName: sheetName
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================
// addDispatchItem — Додавання відправки в аркуш "(відпр)"
// ============================================
function addDispatchItem(payload) {
  try {
    var sheetName = payload.sheetName || '';
    if (!sheetName) {
      return { success: false, error: 'Не вказано аркуш відправки' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }

    var newRow = new Array(DISPATCH_TOTAL_COLS);
    for (var i = 0; i < DISPATCH_TOTAL_COLS; i++) newRow[i] = '';

    newRow[DCOL.NAME] = payload.name || '';
    newRow[DCOL.PARCEL_NUM] = payload.parcelNum || '';
    newRow[DCOL.CITY] = payload.city || '';
    newRow[DCOL.DESCRIPTION] = payload.description || '';
    newRow[DCOL.WEIGHT] = payload.weight || '';
    newRow[DCOL.AMOUNT] = payload.amount || '';
    newRow[DCOL.PAYMENT_TYPE] = payload.paymentType || '';
    newRow[DCOL.CURRENCY] = payload.currency || '';
    newRow[DCOL.ENVELOPE] = payload.envelope || '';
    newRow[DCOL.SENDER_PHONE] = payload.senderPhone || '';
    newRow[DCOL.COMPANY_ID] = payload.companyId || '';

    sheet.appendRow(newRow);
    var newRowNum = sheet.getLastRow();

    writeLog('addDispatchItem', sheetName, newRowNum, 'new',
      'ПІБ: ' + (payload.name || '') + ' | Тел: ' + (payload.senderPhone || '') + ' | Driver UI');

    return {
      success: true,
      sheet: sheetName,
      rowNum: newRowNum
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================
// updateDispatchStatus — Оновлення статусу відправки водієм
// ============================================
function updateDispatchStatus(data) {
  try {
    var sheetName = data.sheetName || '';
    var rowNum = data.rowNum || 0;
    var status = data.status || 'pending';

    if (!sheetName || !rowNum) {
      return { success: false, error: 'Не вказано аркуш або рядок' };
    }

    // Для відпр аркушів статус зберігається лише в логах
    // (відпр аркуші не мають колонки статусу)
    writeLog('updateDispatchStatus', sheetName, rowNum, status,
      'Driver: ' + (data.driverId || '') + ' | Reason: ' + (data.cancelReason || ''));

    return { success: true, rowNum: rowNum, status: status };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================
// editDispatchItem — Редагування відправки в аркуші "(відпр)"
// ============================================
function editDispatchItem(payload) {
  try {
    var sheetName = payload.sheetName || '';
    var rowNum = payload.rowNum || 0;
    if (!sheetName || !rowNum) {
      return { success: false, error: 'Не вказано аркуш або рядок' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }
    if (rowNum > sheet.getLastRow()) {
      return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
    }

    var fieldMap = {
      name: DCOL.NAME,
      parcelNum: DCOL.PARCEL_NUM,
      city: DCOL.CITY,
      description: DCOL.DESCRIPTION,
      weight: DCOL.WEIGHT,
      amount: DCOL.AMOUNT,
      paymentType: DCOL.PAYMENT_TYPE,
      currency: DCOL.CURRENCY,
      envelope: DCOL.ENVELOPE,
      senderPhone: DCOL.SENDER_PHONE
    };

    var fields = payload.fields || {};
    var updated = [];
    for (var key in fields) {
      if (fields.hasOwnProperty(key) && fieldMap.hasOwnProperty(key)) {
        sheet.getRange(rowNum, fieldMap[key] + 1).setValue(fields[key]);
        updated.push(key);
      }
    }

    writeLog('editDispatchItem', sheetName, rowNum, updated.join(', '), JSON.stringify(fields));

    return { success: true, rowNum: rowNum, updated: updated };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================
// editRoutePackage — Редагування посилки в маршрутному аркуші
// ============================================
function editRoutePackage(payload) {
  try {
    var sheetName = payload.sheetName || '';
    var rowNum = payload.rowNum || 0;
    if (!sheetName || !rowNum) {
      return { success: false, error: 'Не вказано аркуш або рядок' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
    }
    if (rowNum > sheet.getLastRow()) {
      return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
    }

    // Використовуємо глобальний FIELD_MAP з усіма 22 полями (ідентичний до Logistics-Cargo)
    var fields = payload.fields || {};
    var updated = [];
    for (var key in fields) {
      if (fields.hasOwnProperty(key) && FIELD_MAP.hasOwnProperty(key)) {
        sheet.getRange(rowNum, FIELD_MAP[key] + 1).setValue(fields[key]);
        updated.push(key);
      }
    }

    writeLog('editRoutePackage', sheetName, rowNum, updated.join(', '), JSON.stringify(fields));

    return { success: true, rowNum: rowNum, updated: updated };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

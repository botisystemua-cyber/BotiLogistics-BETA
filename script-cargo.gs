// ============================================
// BOTI LOGISTICS — CRM ПОСИЛКИ v2.0
// Apps Script API для таблиці "Logistics-Cargo"
// ID: 1E9wYOmVTtlDc52kQAekSpc6rw7Mdnot-m24pRvTUlaY
// ============================================
//
// ІНСТРУКЦІЯ:
// 1. Відкрий таблицю "Бот Посилки" → Розширення → Apps Script
// 2. Видали весь старий код і встав цей файл
// 3. Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Скопіюй URL деплоя
// 5. Встав URL в CRM HTML файл
// ============================================

// ============================================
// КОНФІГУРАЦІЯ
// ============================================

var SPREADSHEET_ID = '1E9wYOmVTtlDc52kQAekSpc6rw7Mdnot-m24pRvTUlaY';

// Один аркуш для всіх посилок (напрямок — в колонці A)
var SHEET_NAME = 'Посилки';
var SHEET_LOGS = 'Логи';

// Порядок колонок (A-W = 23 колонки, індекс 0-22)
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
  PARCEL_STATUS: 10, // K — Статус посилки
  ID: 11,            // L — ІД
  NAME: 12,          // M — ПІБ
  DATE_REG: 13,      // N — дата оформлення
  TIMING: 14,        // O — Таймінг
  SMS_NOTE: 15,      // P — Примітка смс
  DATE_RECEIVE: 16,  // Q — Дата отримання
  PHOTO: 17,         // R — фото
  STATUS: 18,        // S — Статус (CRM: new/work/route/archived/refused/transferred/deleted)
  VEHICLE: 19,       // T — Автомобіль
  COMPANY_ID: 20,    // U — company_id
  DATE_ARCHIVE: 21,  // V — Дата архів
  ARCHIVE_ID: 22     // W — ARCHIVE_ID
};
var TOTAL_COLS = 23;

// Статуси для архівації
var ARCHIVE_STATUSES = ['archived', 'refused', 'deleted', 'transferred'];

// Нормалізація напрямку з таблиці → код CRM
function normalizeDirection(raw) {
  var s = String(raw || '').toLowerCase().trim();
  if (s === 'eu-ua' || s === 'eu→ua' || s === '🇪🇺→🇺🇦') return 'eu-ua';
  if (s === 'ua-eu' || s === 'ua→eu' || s === '🇺🇦→🇪🇺') return 'ua-eu';
  return 'ua-eu'; // default
}

// Код CRM → текст для таблиці
function directionToSheet(dir) {
  return dir === 'eu-ua' ? 'eu-ua' : 'ua-eu';
}

// ============================================
// МАППІНГ полів CRM → індексів колонок
// ============================================
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
  archiveId: COL.ARCHIVE_ID
};

// ============================================
// ГОЛОВНИЙ ОБРОБНИК — doPost
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var companyId = data.companyId || (data.payload && data.payload.companyId) || '';

    switch (action) {
      // --- ЧИТАННЯ ---
      case 'getAll':
        return respond(getAllPackages(companyId));

      case 'getStructure':
        return respond(getStructure());

      // --- СТВОРЕННЯ ---
      case 'addPackage':
        return respond(addPackage(data));

      // --- ОНОВЛЕННЯ ---
      case 'updatePackage':
        return respond(updatePackage(data));

      case 'updateField':
        return respond(updateField(data));

      case 'updateStatus':
        return respond(updateStatus(data));

      case 'bulkUpdateStatus':
        return respond(bulkUpdateStatus(data));

      case 'bulkAssignVehicle':
        return respond(bulkAssignVehicle(data));

      // --- ВИДАЛЕННЯ ---
      case 'deletePackage':
        return respond(deletePackage(data));

      // --- АРХІВАЦІЯ ---
      case 'archivePackage':
        return respond(archivePackage(data));

      case 'bulkArchive':
        return respond(bulkArchive(data));

      // --- ДУБЛІКАТИ ---
      case 'checkDuplicates':
        return respond(checkDuplicates(data));

      default:
        return respond({ success: false, error: 'Невідома дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// doGet — Перевірка здоров'я
// ============================================
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'health';

    switch (action) {
      case 'health':
        return respond({
          success: true,
          version: '2.0',
          service: 'CRM Посилки — BOTI Logistics',
          sheet: SHEET_NAME,
          totalCols: TOTAL_COLS,
          timestamp: new Date().toISOString()
        });

      default:
        return respond({ success: false, error: 'Невідома GET дія: ' + action });
    }
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// ============================================
// getAll — Витягнути ВСІ посилки з обох аркушів
// ============================================
function getAllPackages(companyId) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allPackages = [];

  var sheet = findSheet(ss, SHEET_NAME);
  if (!sheet) {
    return { success: false, error: 'Аркуш "' + SHEET_NAME + '" не знайдено' };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { success: true, packages: [], count: 0, timestamp: new Date().toISOString() };
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS);
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 2;

    if (isEmptyRow(row)) continue;

    var hasIdentity = row[COL.ID] || row[COL.TTN] || row[COL.PHONE] || row[COL.NAME];
    if (!hasIdentity) continue;

    // Фільтр по company_id
    if (companyId) {
      var rowCompanyId = String(row[COL.COMPANY_ID] || '').trim().toLowerCase();
      if (rowCompanyId !== companyId.toLowerCase()) continue;
    }

    var dateReg = formatDate(row[COL.DATE_REG]);

    // Новий лід? (за останні 24 год)
    var isNew24h = false;
    if (dateReg) {
      try {
        var regDate = new Date(dateReg);
        var now = new Date();
        isNew24h = (now.getTime() - regDate.getTime()) < 86400000;
      } catch (e) {}
    }

    var crmStatus = String(row[COL.STATUS] || '').toLowerCase().trim();
    var direction = normalizeDirection(row[COL.DIRECTION]);

    // Визначаємо archiveType зі статусу (refused/deleted/transferred → archiveType)
    var isArchivedStatus = ARCHIVE_STATUSES.indexOf(crmStatus) !== -1;
    var archiveType = isArchivedStatus ? (crmStatus === 'archived' ? 'archived' : crmStatus) : '';

    allPackages.push({
      id: String(row[COL.ID] || ''),
      rowNum: rowNum,
      sheet: SHEET_NAME,

      ttn: String(row[COL.TTN] || ''),
      weight: String(row[COL.WEIGHT] || ''),
      address: String(row[COL.ADDRESS] || ''),
      direction: direction,
      phone: String(row[COL.PHONE] || ''),
      amount: String(row[COL.AMOUNT] || ''),
      payStatus: String(row[COL.PAY_STATUS] || ''),
      payment: String(row[COL.PAYMENT] || ''),
      phoneReg: String(row[COL.PHONE_REG] || ''),
      note: String(row[COL.NOTE] || ''),
      parcelStatus: String(row[COL.PARCEL_STATUS] || ''),
      name: String(row[COL.NAME] || ''),
      dateReg: dateReg,
      timing: String(row[COL.TIMING] || ''),
      smsNote: String(row[COL.SMS_NOTE] || ''),
      dateReceive: formatDate(row[COL.DATE_RECEIVE]),
      photo: String(row[COL.PHOTO] || ''),
      status: crmStatus,
      archiveType: archiveType,
      dateArchive: formatDate(row[COL.DATE_ARCHIVE]),
      archiveId: String(row[COL.ARCHIVE_ID] || ''),
      vehicle: String(row[COL.VEHICLE] || ''),

      isNew: isNew24h,
      isArchived: isArchivedStatus
    });
  }

  return {
    success: true,
    packages: allPackages,
    count: allPackages.length,
    timestamp: new Date().toISOString()
  };
}

// ============================================
// addPackage — Додати нову посилку
// + перевірка дублікатів перед додаванням
// ============================================
function addPackage(data) {
  var fields = data.fields;
  if (!fields) {
    return { success: false, error: 'Відсутні поля (fields)' };
  }

  var direction = fields.direction || 'ua-eu';

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, SHEET_NAME);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + SHEET_NAME };
  }

  // --- ПЕРЕВІРКА ДУБЛІКАТІВ ---
  var duplicates = [];
  var checkTTN = fields.ttn ? String(fields.ttn).trim() : '';
  var checkPhone = fields.phone ? String(fields.phone).trim() : '';
  var checkId = fields.id ? String(fields.id).trim() : '';

  if (checkTTN || checkPhone || checkId) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var chkValues = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
      for (var r = 0; r < chkValues.length; r++) {
        var existRow = chkValues[r];
        var existStatus = String(existRow[COL.STATUS] || '').toLowerCase().trim();
        if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

        var existTTN = String(existRow[COL.TTN] || '').trim();
        var existPhone = String(existRow[COL.PHONE] || '').trim();
        var existId = String(existRow[COL.ID] || '').trim();

        var isDuplicate = false;
        var reason = '';

        if (checkId && existId && checkId === existId) {
          isDuplicate = true;
          reason = 'ІД: ' + checkId;
        }
        if (checkTTN && existTTN && checkTTN === existTTN) {
          isDuplicate = true;
          reason = 'ТТН: ' + checkTTN;
        }
        if (!isDuplicate && checkPhone && existPhone && checkPhone === existPhone) {
          var checkName = String(fields.name || '').trim().toLowerCase();
          var existName = String(existRow[COL.NAME] || '').trim().toLowerCase();
          if (checkName && existName && checkName === existName) {
            isDuplicate = true;
            reason = 'Телефон+ПіБ: ' + checkPhone;
          }
        }

        if (isDuplicate) {
          duplicates.push({
            sheet: SHEET_NAME,
            rowNum: r + 2,
            ttn: existTTN,
            phone: existPhone,
            name: String(existRow[COL.NAME] || ''),
            status: existStatus,
            reason: reason
          });
        }
      }
    }
  }

  if (duplicates.length > 0 && !data.force) {
    writeLog('addPackage:DUPLICATE', SHEET_NAME, 0, 'blocked',
      'Знайдено ' + duplicates.length + ' дублікатів | ' + duplicates[0].reason);

    return {
      success: false,
      error: 'duplicate',
      message: 'Знайдено ' + duplicates.length + ' можливих дублікатів',
      duplicates: duplicates
    };
  }

  // --- СТВОРЕННЯ РЯДКА ---
  var newRow = new Array(TOTAL_COLS);
  for (var i = 0; i < TOTAL_COLS; i++) {
    newRow[i] = '';
  }

  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      newRow[FIELD_MAP[field]] = fields[field];
    }
  }

  // Автозаповнення
  if (!newRow[COL.DATE_REG]) {
    newRow[COL.DATE_REG] = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  }
  if (!newRow[COL.ID]) {
    newRow[COL.ID] = 'crm_' + new Date().getTime();
  }
  // Напрямок завжди записуємо в колонку A
  newRow[COL.DIRECTION] = directionToSheet(direction);
  if (!newRow[COL.STATUS]) {
    newRow[COL.STATUS] = 'new';
  }
  if (data.companyId) {
    newRow[COL.COMPANY_ID] = data.companyId;
  }

  sheet.appendRow(newRow);
  var newRowNum = sheet.getLastRow();

  writeLog('addPackage', SHEET_NAME, newRowNum, 'new',
    'ПіБ: ' + (fields.name || '') + ' | ТТН: ' + (fields.ttn || '') + ' | Тел: ' + (fields.phone || '') +
    (duplicates.length > 0 ? ' | FORCE (дублікат ігноровано)' : ''));

  return {
    success: true,
    sheet: SHEET_NAME,
    rowNum: newRowNum,
    id: newRow[COL.ID],
    direction: direction,
    duplicatesIgnored: duplicates.length
  };
}

// ============================================
// updatePackage — Оновити посилку
// + верифікація по ІД щоб не перезаписати чужий рядок
// ============================================
function updatePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var fields = data.fields;

  if (!sheetName || !rowNum || !fields) {
    return { success: false, error: 'Відсутні sheet, rowNum або fields' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує (lastRow: ' + sheet.getLastRow() + ')' };
  }

  // --- ВЕРИФІКАЦІЯ ПО ІД ---
  // Якщо передано expectedId — перевіряємо що рядок не змінився
  if (data.expectedId) {
    var currentId = String(sheet.getRange(rowNum, COL.ID + 1).getValue() || '').trim();
    if (currentId !== String(data.expectedId).trim()) {
      writeLog('updatePackage:CONFLICT', sheetName, rowNum, 'blocked',
        'Очікувався ІД: ' + data.expectedId + ', фактичний: ' + currentId);

      return {
        success: false,
        error: 'conflict',
        message: 'Рядок змінився (можливо видалено/переміщено). Очікувався ІД: ' + data.expectedId + ', фактичний: ' + currentId
      };
    }
  }

  // --- ПЕРЕВІРКА ЧИ НЕ АРХІВОВАНИЙ ---
  var currentStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();
  if (ARCHIVE_STATUSES.indexOf(currentStatus) !== -1 && !data.force) {
    return {
      success: false,
      error: 'archived',
      message: 'Запис вже архівований (статус: ' + currentStatus + '). Використайте force=true для перезапису'
    };
  }

  var updated = [];
  for (var field in fields) {
    if (fields.hasOwnProperty(field) && FIELD_MAP.hasOwnProperty(field)) {
      sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(fields[field]);
      updated.push(field);
    }
  }

  writeLog('updatePackage', sheetName, rowNum, updated.join(', '), JSON.stringify(fields));

  return { success: true, updated: updated, sheet: sheetName, rowNum: rowNum };
}

// ============================================
// updateField — Оновити одне поле
// ============================================
function updateField(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var field = data.field;
  var value = data.value;

  if (!sheetName || !rowNum || !field) {
    return { success: false, error: 'Відсутні sheet, rowNum або field' };
  }

  if (!FIELD_MAP.hasOwnProperty(field)) {
    return { success: false, error: 'Невідоме поле: ' + field };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  sheet.getRange(rowNum, FIELD_MAP[field] + 1).setValue(value);

  writeLog('updateField', sheetName, rowNum, field, String(value));

  return { success: true, sheet: sheetName, rowNum: rowNum, field: field };
}

// ============================================
// updateStatus — Змінити CRM статус
// + автоматичне логування дати архіву
// ============================================
function updateStatus(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var newStatus = data.status;

  if (!sheetName || !rowNum || !newStatus) {
    return { success: false, error: 'Відсутні sheet, rowNum або status' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  var oldStatus = String(sheet.getRange(rowNum, COL.STATUS + 1).getValue() || '').toLowerCase().trim();

  // Ставимо новий статус
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(newStatus);

  // Якщо архівний статус — ставимо дату
  if (ARCHIVE_STATUSES.indexOf(newStatus) !== -1) {
    sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
      Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
    );
  }

  writeLog('updateStatus', sheetName, rowNum, oldStatus + ' → ' + newStatus, '');

  return { success: true, sheet: sheetName, rowNum: rowNum, status: newStatus, oldStatus: oldStatus };
}

// ============================================
// bulkUpdateStatus — Масова зміна статусу
// ============================================
function bulkUpdateStatus(data) {
  var items = data.items;
  var newStatus = data.status;

  if (!items || !items.length || !newStatus) {
    return { success: false, error: 'Відсутні items або status' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var needDate = ARCHIVE_STATUSES.indexOf(newStatus) !== -1;
  var count = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = findSheet(ss, item.sheet);
    if (!sheet) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Аркуш не знайдено' });
      continue;
    }
    if (item.rowNum > sheet.getLastRow()) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Рядок не існує' });
      continue;
    }

    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue(newStatus);
    if (needDate) {
      sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
    } else {
      // Відновлення: очищаємо архівні поля
      sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue('');
      sheet.getRange(item.rowNum, COL.ARCHIVE_ID + 1).setValue('');
    }
    count++;
  }

  writeLog('bulkUpdateStatus', 'bulk', 0, newStatus, count + '/' + items.length + ' оновлено');

  return {
    success: true,
    count: count,
    total: items.length,
    status: newStatus,
    errors: errors.length > 0 ? errors : undefined
  };
}

// ============================================
// deletePackage — Видалити (статус = deleted + дата)
// ============================================
function deletePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Відсутні sheet або rowNum' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  // Зберігаємо інфо для логу
  var recordId = String(sheet.getRange(rowNum, COL.ID + 1).getValue() || '');
  var recordName = String(sheet.getRange(rowNum, COL.NAME + 1).getValue() || '');

  // Позначаємо як видалений
  sheet.getRange(rowNum, COL.STATUS + 1).setValue('deleted');
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(
    Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd')
  );

  writeLog('deletePackage', sheetName, rowNum, 'deleted',
    'ІД: ' + recordId + ' | ПіБ: ' + recordName);

  return { success: true, sheet: sheetName, rowNum: rowNum, id: recordId };
}

// ============================================
// bulkAssignVehicle — Масове призначення авто
// ============================================
function bulkAssignVehicle(data) {
  var items = data.items;
  var vehicle = data.vehicle;

  if (!items || !items.length || !vehicle) {
    return { success: false, error: 'Відсутні items або vehicle' };
  }

  writeLog('bulkAssignVehicle', 'bulk', 0, vehicle, items.length + ' items');

  return { success: true, count: items.length, vehicle: vehicle };
}

// ============================================
// archivePackage — Архівувати одну посилку
// In-place: оновлює STATUS, DATE_ARCHIVE, ARCHIVE_ID
// без копіювання в зовнішню таблицю
// ============================================
function archivePackage(data) {
  var sheetName = data.sheet;
  var rowNum = data.rowNum;
  var user = data.user || 'crm';
  var reason = data.reason || 'manual';

  if (!sheetName || !rowNum) {
    return { success: false, error: 'Відсутні sheet або rowNum' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheet(ss, sheetName);
  if (!sheet) {
    return { success: false, error: 'Аркуш не знайдено: ' + sheetName };
  }

  if (rowNum > sheet.getLastRow()) {
    return { success: false, error: 'Рядок ' + rowNum + ' не існує' };
  }

  // Читаємо поточний рядок
  var rowData = sheet.getRange(rowNum, 1, 1, TOTAL_COLS).getValues()[0];

  // Перевірка: чи не вже архівований
  var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
  if (existingArchiveId) {
    return {
      success: false,
      error: 'already_archived',
      message: 'Запис вже архівований: ' + existingArchiveId
    };
  }

  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var archiveId = generateArchiveId_();

  // In-place: оновлюємо 3 колонки в тій самій таблиці
  // Зберігаємо конкретний тип архівації (refused/deleted/transferred/archived)
  var archiveStatus = (reason && ARCHIVE_STATUSES.indexOf(reason) !== -1) ? reason : 'archived';
  sheet.getRange(rowNum, COL.STATUS + 1).setValue(archiveStatus);
  sheet.getRange(rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
  sheet.getRange(rowNum, COL.ARCHIVE_ID + 1).setValue(archiveId);

  var recordId = String(rowData[COL.ID] || '');
  writeLog('archivePackage', sheetName, rowNum, 'archived',
    'ІД: ' + recordId + ' | ArchiveID: ' + archiveId + ' | by: ' + user + ' | reason: ' + reason);

  return {
    success: true,
    sheet: sheetName,
    rowNum: rowNum,
    id: recordId,
    archiveId: archiveId
  };
}

// ============================================
// bulkArchive — Масова архівація
// In-place: оновлює STATUS, DATE_ARCHIVE, ARCHIVE_ID
// без копіювання в зовнішню таблицю
// ============================================
function bulkArchive(data) {
  var items = data.items; // масив { sheet, rowNum }
  var user = data.user || 'crm';
  var reason = data.reason || 'bulk';

  if (!items || !items.length) {
    return { success: false, error: 'Відсутні items' };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dateNow = Utilities.formatDate(new Date(), 'Europe/Kiev', 'yyyy-MM-dd');
  var count = 0;
  var errors = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sheet = findSheet(ss, item.sheet);
    if (!sheet) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Аркуш не знайдено' });
      continue;
    }
    if (item.rowNum > sheet.getLastRow()) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Рядок не існує' });
      continue;
    }

    var rowData = sheet.getRange(item.rowNum, 1, 1, TOTAL_COLS).getValues()[0];
    var existingArchiveId = String(rowData[COL.ARCHIVE_ID] || '').trim();
    if (existingArchiveId) {
      errors.push({ sheet: item.sheet, rowNum: item.rowNum, error: 'Вже архівовано' });
      continue;
    }

    var archiveId = generateArchiveId_();

    // In-place: оновлюємо 3 колонки
    sheet.getRange(item.rowNum, COL.STATUS + 1).setValue('archived');
    sheet.getRange(item.rowNum, COL.DATE_ARCHIVE + 1).setValue(dateNow);
    sheet.getRange(item.rowNum, COL.ARCHIVE_ID + 1).setValue(archiveId);
    count++;
  }

  writeLog('bulkArchive', 'bulk', 0, 'archived',
    count + '/' + items.length + ' архівовано in-place | by: ' + user + ' | reason: ' + reason);

  return {
    success: true,
    count: count,
    total: items.length,
    errors: errors.length > 0 ? errors : undefined
  };
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
// checkDuplicates — Перевірити дублікати
// ============================================
// Параметри:
//   ttn: номер ТТН
//   phone: телефон
//   id: ІД запису
//   name: ПіБ (для м'якої перевірки з телефоном)
//   excludeRow: { sheet, rowNum } — виключити цей рядок
// ============================================
function checkDuplicates(data) {
  var checkTTN = data.ttn ? String(data.ttn).trim() : '';
  var checkPhone = data.phone ? String(data.phone).trim() : '';
  var checkId = data.id ? String(data.id).trim() : '';
  var checkName = data.name ? String(data.name).trim().toLowerCase() : '';
  var excludeSheet = data.excludeRow ? data.excludeRow.sheet : '';
  var excludeRowNum = data.excludeRow ? data.excludeRow.rowNum : 0;

  if (!checkTTN && !checkPhone && !checkId) {
    return { success: true, duplicates: [], count: 0 };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var duplicates = [];

  var chkSheet = findSheet(ss, SHEET_NAME);
  if (chkSheet) {
    var lastRow = chkSheet.getLastRow();
    if (lastRow >= 2) {
      var values = chkSheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();

      for (var r = 0; r < values.length; r++) {
        var rowNum = r + 2;
        var row = values[r];

        if (SHEET_NAME === excludeSheet && rowNum === excludeRowNum) continue;
        if (isEmptyRow(row)) continue;

        var existStatus = String(row[COL.STATUS] || '').toLowerCase().trim();
        if (ARCHIVE_STATUSES.indexOf(existStatus) !== -1) continue;

        var existTTN = String(row[COL.TTN] || '').trim();
        var existPhone = String(row[COL.PHONE] || '').trim();
        var existId = String(row[COL.ID] || '').trim();
        var existName = String(row[COL.NAME] || '').trim().toLowerCase();

        var matchReasons = [];

        if (checkId && existId && checkId === existId) {
          matchReasons.push('ІД');
        }
        if (checkTTN && existTTN && checkTTN === existTTN) {
          matchReasons.push('ТТН');
        }
        if (checkPhone && existPhone && checkPhone === existPhone && checkName && existName && checkName === existName) {
          matchReasons.push('Телефон+ПіБ');
        }

        if (matchReasons.length > 0) {
          duplicates.push({
            sheet: SHEET_NAME,
            rowNum: rowNum,
            id: existId,
            ttn: existTTN,
            phone: existPhone,
            name: String(row[COL.NAME] || ''),
            status: existStatus,
            dateReg: formatDate(row[COL.DATE_REG]),
            matchReasons: matchReasons
          });
        }
      }
    }
  }

  return {
    success: true,
    duplicates: duplicates,
    count: duplicates.length
  };
}

// ============================================
// getStructure — Структура таблиці (дебаг)
// ============================================
function getStructure() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var result = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    var sample = lastRow > 1
      ? sheet.getRange(2, 1, Math.min(2, lastRow - 1), lastCol).getValues()
      : [];

    result.push({
      sheet: name,
      rows: lastRow,
      cols: lastCol,
      headers: headers,
      sample: sample
    });
  }

  return { success: true, sheets: result };
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
// ДОПОМІЖНІ ФУНКЦІЇ
// ============================================

// Знайти аркуш (з fallback для старої назви)
function findSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  // Backward compat: якщо шукають "Посилки" але аркуш ще називається "Реєстрація ТТН"
  if (name === 'Посилки') {
    sheet = ss.getSheetByName('Реєстрація ТТН');
    if (sheet) return sheet;
  }
  return null;
}

// Перевірка порожнього рядка
function isEmptyRow(row) {
  for (var c = 0; c < row.length && c < TOTAL_COLS; c++) {
    if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
      return false;
    }
  }
  return true;
}

// Форматування дати → YYYY-MM-DD
function formatDate(value) {
  if (!value) return '';

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return Utilities.formatDate(value, 'Europe/Kiev', 'yyyy-MM-dd');
  }

  var str = String(value).trim();
  if (!str) return '';

  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.substring(0, 10);

  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(str)) {
    var p = str.split('.');
    return p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
  }

  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
    var p2 = str.split('/');
    return p2[2] + '-' + ('0' + p2[1]).slice(-2) + '-' + ('0' + p2[0]).slice(-2);
  }

  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Europe/Kiev', 'yyyy-MM-dd');
    }
  } catch (e) {}

  return '';
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
  ui.createMenu('CRM Посилки')
    .addItem('Структура таблиці', 'testStructure')
    .addItem('Тест: всі посилки', 'testGetAll')
    .addItem('Тест: знайти аркуші', 'testFindSheets')
    .addToUi();
}

// ============================================
// ТЕСТИ
// ============================================

function testFindSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();

  Logger.log('=== ВСІ АРКУШІ ===');
  for (var i = 0; i < sheets.length; i++) {
    Logger.log('  [' + i + '] "' + sheets[i].getName() + '" (' + sheets[i].getLastRow() + ' рядків)');
  }

  Logger.log('');
  var mainSheet = findSheet(ss, SHEET_NAME);
  Logger.log('Посилки: ' + (mainSheet ? 'ЗНАЙДЕНО (' + mainSheet.getLastRow() + ' рядків)' : 'НЕ ЗНАЙДЕНО'));
}

function testGetAll() {
  var result = getAllPackages();
  Logger.log('Всього посилок: ' + result.count);

  if (result.packages && result.packages.length > 0) {
    for (var i = 0; i < Math.min(5, result.packages.length); i++) {
      var p = result.packages[i];
      Logger.log(
        '[' + p.sheet + ' #' + p.rowNum + '] ' +
        'ПіБ: ' + (p.name || '-') +
        ' | ТТН: ' + (p.ttn || '-') +
        ' | Статус: ' + (p.status || '(пусто)') +
        ' | ArchiveID: ' + (p.archiveId || '-')
      );
    }
  }
}

function testStructure() {
  var result = getStructure();
  Logger.log('=== СТРУКТУРА ===');
  for (var i = 0; i < result.sheets.length; i++) {
    var s = result.sheets[i];
    Logger.log('[' + s.sheet + '] ' + s.rows + ' рядків, ' + s.cols + ' колонок');
    Logger.log('  Колонки: ' + s.headers.join(' | '));
  }
}


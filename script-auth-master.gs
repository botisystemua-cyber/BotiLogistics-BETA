// ═══════════════════════════════════════════════════════
// BOTI-LOGISTICS — СКРИПТ АВТЕНТИФІКАЦІЇ (МАЙСТЕР-ТАБЛИЦЯ)
// Прив'язати до Google Sheet "Logistics-паролі"
// Розгорнути як Web App (доступ: Anyone)
// ═══════════════════════════════════════════════════════
//
// Структура таблиці "Паролі" (аркуш):
//   A: Компанія
//   B: Роль (Менеджер / Власник)
//   C: Email
//   D: Пароль
//   E: company_id (текстовий ключ компанії, напр. "yura", "demo")
//

var SHEET_NAME = 'Паролі';

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    if (action === 'login')          return out(handleLogin(payload));
    if (action === 'getUsers')       return out(handleGetUsers(payload));
    if (action === 'addUser')        return out(handleAddUser(payload));
    if (action === 'deleteUser')     return out(handleDeleteUser(payload));
    if (action === 'toggleStatus')   return out(handleToggleStatus(payload));
    if (action === 'changePassword') return out(handleChangePassword(payload));

    return out({ success: false, error: 'Невідома дія: ' + action });
  } catch (err) {
    return out({ success: false, error: err.message });
  }
}

function doGet(e) {
  return out({ success: false, error: 'Використовуйте POST' });
}

// ═══ HELPERS ═══

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getAllRows() {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  // Пропускаємо заголовок (рядок 1)
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][2]) continue; // пропускаємо рядки без email
    rows.push({
      row: i + 1, // номер рядка в таблиці (1-based)
      company:  String(data[i][0]).trim(),
      role:     String(data[i][1]).trim(),
      email:    String(data[i][2]).trim().toLowerCase(),
      password: String(data[i][3]).trim(),
      companyId: String(data[i][4]).trim()
    });
  }
  return rows;
}

function findByEmail(email) {
  var rows = getAllRows();
  email = email.toLowerCase().trim();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].email === email) return rows[i];
  }
  return null;
}

// ═══ ЛОГІН ═══

function handleLogin(p) {
  var email = (p.email || '').trim().toLowerCase();
  var password = (p.password || '').trim();
  var expectedRole = (p.role || '').trim();

  if (!email || !password) {
    return { success: false, error: 'Введіть email та пароль' };
  }

  var user = findByEmail(email);

  if (!user) {
    return { success: false, error: 'Користувача з таким email не знайдено' };
  }

  if (user.password !== password) {
    return { success: false, error: 'Невірний пароль' };
  }

  // Перевіряємо роль якщо вказана (менеджер натиснув на картку менеджера)
  if (expectedRole && user.role !== expectedRole) {
    return { success: false, error: 'Ваша роль: ' + user.role + '. Оберіть правильний вхід.' };
  }

  return {
    success: true,
    name: user.email, // ім'я = email
    role: user.role,
    companyName: user.company,
    companyId: user.companyId
  };
}

// ═══ ОТРИМАТИ КОРИСТУВАЧІВ (тільки для Власника) ═══

function handleGetUsers(p) {
  var owner = verifyOwner(p.ownerEmail, p.ownerPassword);
  if (!owner) return { success: false, error: 'Немає доступу' };

  var rows = getAllRows();
  // Показуємо тільки користувачів тієї ж компанії
  var users = [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].company === owner.company) {
      users.push({
        name: rows[i].email,
        role: rows[i].role,
        status: 'Активний'
      });
    }
  }

  return { success: true, users: users };
}

// ═══ ДОДАТИ КОРИСТУВАЧА ═══

function handleAddUser(p) {
  var owner = verifyOwner(p.ownerEmail, p.ownerPassword);
  if (!owner) return { success: false, error: 'Немає доступу' };

  var newEmail = (p.newEmail || '').trim().toLowerCase();
  var newPassword = (p.newPassword || '').trim();
  var newRole = (p.newRole || 'Менеджер').trim();

  if (!newEmail) return { success: false, error: 'Введіть email' };
  if (!newPassword || newPassword.length < 4) return { success: false, error: 'Пароль мінімум 4 символи' };

  // Перевірка чи email вже існує
  var existing = findByEmail(newEmail);
  if (existing) return { success: false, error: 'Користувач з таким email вже існує' };

  // Додаємо в таблицю — та ж компанія і company_id як у власника
  var sheet = getSheet();
  sheet.appendRow([owner.company, newRole, newEmail, newPassword, owner.companyId]);

  return { success: true, message: 'Користувача ' + newEmail + ' додано (' + newRole + ')' };
}

// ═══ ВИДАЛИТИ КОРИСТУВАЧА ═══

function handleDeleteUser(p) {
  var owner = verifyOwner(p.ownerEmail, p.ownerPassword);
  if (!owner) return { success: false, error: 'Немає доступу' };

  var targetEmail = (p.targetEmail || '').trim().toLowerCase();
  var target = findByEmail(targetEmail);

  if (!target) return { success: false, error: 'Користувача не знайдено' };
  if (target.company !== owner.company) return { success: false, error: 'Немає доступу' };

  // Якщо видаляємо Власника — потрібен його пароль
  if (target.role === 'Власник') {
    var targetPassword = (p.targetPassword || '').trim();
    if (target.password !== targetPassword) {
      return { success: false, error: 'Невірний пароль власника' };
    }
  }

  var sheet = getSheet();
  sheet.deleteRow(target.row);

  return { success: true, message: 'Видалено' };
}

// ═══ БЛОКУВАННЯ / РОЗБЛОКУВАННЯ ═══

function handleToggleStatus(p) {
  // Поки що спрощена версія — просто повертаємо успіх
  // В майбутньому можна додати колонку "Статус" в таблицю
  var owner = verifyOwner(p.ownerEmail, p.ownerPassword);
  if (!owner) return { success: false, error: 'Немає доступу' };

  return { success: true, message: 'Статус змінено' };
}

// ═══ ЗМІНИТИ ПАРОЛЬ ═══

function handleChangePassword(p) {
  var email = (p.email || '').trim().toLowerCase();
  var oldPassword = (p.oldPassword || '').trim();
  var newPassword = (p.newPassword || '').trim();

  if (!newPassword || newPassword.length < 4) {
    return { success: false, error: 'Новий пароль мінімум 4 символи' };
  }

  var user = findByEmail(email);
  if (!user) return { success: false, error: 'Користувача не знайдено' };
  if (user.password !== oldPassword) return { success: false, error: 'Невірний старий пароль' };

  // Оновлюємо пароль (колонка D = індекс 4)
  var sheet = getSheet();
  sheet.getRange(user.row, 4).setValue(newPassword);

  return { success: true, message: 'Пароль змінено' };
}

// ═══ ПЕРЕВІРКА ВЛАСНИКА ═══

function verifyOwner(email, password) {
  if (!email || !password) return null;
  var user = findByEmail(email);
  if (!user) return null;
  if (user.password !== password) return null;
  if (user.role !== 'Власник') return null;
  return user;
}

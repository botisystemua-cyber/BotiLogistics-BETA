# BOTI Logistics — CRM System

CRM-система для логістичної компанії BOTI. Управління пасажирами, посилками (карго), маршрутами та водіями.

Архітектура: статичні HTML-сторінки (фронтенд) + Google Apps Script (бекенд) + Google Sheets (база даних).

---

## Структура проекту

```
BOTI-Logistics/
├── Login.html                    — Авторизація (логін/пароль → перенаправлення за роллю)
├── Passengers.html               — CRM пасажирів (UA↔EU)
├── Cargo.html                    — CRM посилок (ТТН, кур'єр, маршрути карго)
├── Drivers.html                  — Панель водія (маршрути, навігація, статуси)
├── script-passengers.gs          — Бекенд: CRM пасажирів
├── script-cargo.gs               — Бекенд: CRM посилок
├── script-marshrut-passengers.gs — Бекенд: маршрути пасажирів
├── script-marshrut-cargo.gs      — Бекенд: маршрути посилок
├── script-auth-master.gs         — Бекенд: авторизація
└── README.md
```

---

## Ролі користувачів

| Роль | Сторінка | Можливості |
|------|----------|------------|
| Менеджер | `Passengers.html` | CRUD пасажирів, створення маршрутів, архівація, розсилка |
| Менеджер | `Cargo.html` | CRUD посилок, маршрути карго, архівація, розсилка |
| Водій | `Drivers.html` | Перегляд маршрутів (пасажири + карго), оновлення статусів, навігація |

Авторизація через `Login.html` → перенаправлення на відповідну сторінку залежно від ролі.

---

## API Endpoints (Google Apps Script deploy URLs)

### Авторизація
```
URL:    https://script.google.com/macros/s/AKfycbw9EZi_03T0nvJ5WdJzl9rfHImjE6A4kkt4-pVqQmpE452vciAzD0tiF62AFbPJBYzkLw/exec
Скрипт: script-auth-master.gs
Де:     Login.html → AUTH_API_URL
```

### CRM Пасажири
```
URL:    https://script.google.com/macros/s/AKfycbwr_CgMtGzKdMj0H5JRp2Y3xEaVjYg78s5A9pT_qfsysEVQrl0FBSNJiEB1WiFx2pIW6Q/exec
Скрипт: script-passengers.gs
Де:     Passengers.html → API_URL
```

### Маршрути Пасажирів
```
URL:    https://script.google.com/macros/s/AKfycbxWaPrmjKqNdAAZib73hG-aJXQYXyJC-lxwxFQS15pgFp7Xhpw6Ibcb3nnbRB38y98S/exec
Скрипт: script-marshrut-passengers.gs
Де:     Passengers.html → ROUTE_API_URL
        Drivers.html → CONFIG.PASSENGER_API_URL
```

### CRM Посилки (Карго)
```
URL:    https://script.google.com/macros/s/AKfycbwNTUVFhk4nD4X7LB_JIc6rUXkjgj6w0dhHLHTCE73PV-JczFiGTM8ZyB_QtqBiOg1ugw/exec
Скрипт: script-cargo.gs
Де:     Cargo.html → API_URL
```

### Маршрути Посилок (Карго)
```
URL:    https://script.google.com/macros/s/AKfycbzjP4grKvDrqUPq5Y_NqX5nMMJPKncggWfbTKjI_-cPKEv_aE9xuxZaW2PW8RrJsAqm/exec
Скрипт: script-marshrut-cargo.gs
Де:     Cargo.html → ROUTE_API_URL
        Drivers.html → CONFIG.DELIVERY_API_URL
```

### Drivers.html (водій) — використовує 2 API
```
DELIVERY_API_URL:  https://script.google.com/macros/s/AKfycbzjP4grKvDrqUPq5Y_NqX5nMMJPKncggWfbTKjI_-cPKEv_aE9xuxZaW2PW8RrJsAqm/exec  (маршрути карго)
PASSENGER_API_URL: https://script.google.com/macros/s/AKfycbxWaPrmjKqNdAAZib73hG-aJXQYXyJC-lxwxFQS15pgFp7Xhpw6Ibcb3nnbRB38y98S/exec  (маршрути пасажирів)
```

---

## Google Sheets (база даних)

Кожен бекенд-скрипт працює зі своєю Google Таблицею.

### Пасажири — `script-passengers.gs`
- **Spreadsheet ID:** `1SvWAVYNKkWl7Sx_wWhTWPlDyKU9hIQCQzcWBTUGG2i8`
- **Route Spreadsheet ID:** `1fYO1ClIP26S4xYgcsT_0LVCWVrqkAL5MkehXvL-Yni0`
- Аркуші: `Україна-єв` (UA→EU), `Європа-ук` (EU→UA), `Логи`

### Посилки (Карго) — `script-cargo.gs`
- **Spreadsheet ID:** `1E9wYOmVTtlDc52kQAekSpc6rw7Mdnot-m24pRvTUlaY`
- Аркуші: `Реєстрація ТТН` (UA→EU), `Виклик курєра` (EU→UA), `Логи`

### Маршрути Пасажирів — `script-marshrut-passengers.gs`
- **Spreadsheet ID:** `1fYO1ClIP26S4xYgcsT_0LVCWVrqkAL5MkehXvL-Yni0`
- Аркуші: динамічні (створюються через CRM), `Логи`, `Провірка розсилки`

### Маршрути Посилок (Карго) — `script-marshrut-cargo.gs`
- **Spreadsheet ID:** `17g3TFYg11EqdQ9eGrOKQV3n_nqPDFx7dqsJVaGWeDOo`
- Аркуші: `Братислава марш.`, `Нітра марш.`, `Словаччина марш.`, `Кошице+прешов марш.` та інші
- Спец. аркуші: `Маршрути водіїв` (логи), `Провірка розсилки` (розсилка)

### Логи (централізовані)
- **Spreadsheet ID:** `1MxX6aA1kZYmBwgI2g2pylxSZz-Tzvi0v0YyJENWXBPw`
- **Аркуш:** `Логи`
- Всі скрипти пишуть в одну таблицю
- Колонки: `Дата/Час`, `Модуль`, `Користувач`, `Дія`, `Об'єкт`, `К-сть`, `Деталі`

| LOG_SOURCE | Скрипт |
|------------|--------|
| `CRM-Посилки` | `script-cargo.gs` |
| `CRM-Пасажири` | `script-passengers.gs` |
| `Маршрути-Посилки` | `script-marshrut-cargo.gs` |
| `Маршрути-Пасажири` | `script-marshrut-passengers.gs` |

---

## Як працює система

### Потік даних
1. **Менеджер** створює запис (пасажир/посилка) через CRM-сторінку
2. Дані зберігаються в Google Sheets через Apps Script API
3. **Менеджер** формує маршрут (відправку) — дані копіюються на відповідний аркуш маршруту
4. **Водій** бачить маршрут у `Drivers.html`, оновлює статуси (забрав, доставив)
5. Після завершення — записи архівуються

### Видалення рядків (safe delete)
При видаленні записів з аркушів маршрутів використовується захист від помилки Google Sheets "не може мати 0 рядків":
- Перед `deleteRow()` перевіряється чи це не останній рядок даних
- Якщо останній — `clearContent()` замість видалення
- Додатково `try-catch` fallback на всі операції видалення

### Архівація
Архівація працює через статуси (`archived`, `refused`, `deleted`, `transferred`) і поле `ARCHIVE_ID` / `Дата архів`. Записи не видаляються фізично, а позначаються відповідним статусом.

---

## Деплой скриптів

1. Відкрий відповідну Google Таблицю → **Розширення** → **Apps Script**
2. Заміни код на відповідний `.gs` файл з цього репозиторію
3. **Deploy** → **New deployment** → **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Скопіюй новий URL і оновити відповідну змінну в HTML файлі
5. Закоміть зміни в репозиторій

### Де оновлювати URL після деплою

| Скрипт | HTML файл | Змінна |
|--------|-----------|--------|
| `script-auth-master.gs` | `Login.html` | `AUTH_API_URL` |
| `script-passengers.gs` | `Passengers.html` | `API_URL` |
| `script-marshrut-passengers.gs` | `Passengers.html` | `ROUTE_API_URL` |
| `script-marshrut-passengers.gs` | `Drivers.html` | `CONFIG.PASSENGER_API_URL` |
| `script-cargo.gs` | `Cargo.html` | `API_URL` |
| `script-marshrut-cargo.gs` | `Cargo.html` | `ROUTE_API_URL` |
| `script-marshrut-cargo.gs` | `Drivers.html` | `CONFIG.DELIVERY_API_URL` |

---

## Відомі баги / нотатки

- **Drivers.html — пасажирські маршрути**: у пасажирів немає функції "Відправки" (dispatch). Фільтр "Відправки" та кнопка "Додати → Відправки" стосуються тільки карго маршрутів. Кнопка навігації в карточці пасажира — "Карта" (пункт призначення), а не "Відправка".

# BOTI Logistics — CRM System

CRM-система для логістичної компанії. Управління пасажирами, посилками, маршрутами та водіями.

## Структура проекту

### HTML-фронтенд (статичні сторінки)

| Файл | Призначення |
|------|-------------|
| `Login.html` | Авторизація (логін/пароль → перенаправлення за роллю) |
| `Passengers.html` | CRM пасажирів (UA↔EU) — менеджер |
| `Cargo.html` | CRM посилок (ТТН, кур'єр) — менеджер |
| `Drivers.html` | Панель водія — маршрути, навігація, статуси |

### Google Apps Script (бекенд)

| Файл | Сервіс | Таблиця (Spreadsheet ID) |
|------|--------|--------------------------|
| `script-passangers.gs` | CRM Пасажири | `1SvWAVYNKkWl7Sx_wWhTWPlDyKU9hIQCQzcWBTUGG2i8` |
| `script-cargo.gs` | CRM Посилки | `1E9wYOmVTtlDc52kQAekSpc6rw7Mdnot-m24pRvTUlaY` |
| `script-marshrut-passengers.gs` | Маршрути Пасажири | `1fYO1ClIP26S4xYgcsT_0LVCWVrqkAL5MkehXvL-Yni0` |
| `script-marshrut-cargo.gs` | Маршрути Посилки | `17g3TFYg11EqdQ9eGrOKQV3n_nqPDFx7dqsJVaGWeDOo` |
| `script-auth-master.gs` | Авторизація | Окрема таблиця з аркушем "Паролі" |
| `skript-arhiv.gs` | Архів | `1Kmf6NF1sJUi-j3SamrhUqz337pcZSvZCUkGxBzari6U` |

## API URL-и (deploy endpoints)

### Авторизація
```
AUTH_API_URL = AKfycbw9EZi_03T0nvJ5WdJzl9rfHImjE6A4kkt4-pVqQmpE452vciAzD0tiF62AFbPJBYzkLw
Скрипт: script-auth-master.gs
```

### Пасажири (CRM)
```
API_URL = AKfycbw1gPO3mpGK4uAxnqWbnGok_2OUzpYf8HrHGQ7SmS2pZtuE6dyi3z_oOCX8jssf8_ZWhg
Скрипт: script-passangers.gs
Таблиця: Logistics-Passengers
```

### Посилки (CRM)
```
API_URL = AKfycbyjygnhQtrTFrbmmuNEdulvsIgD7eOKagyD7_8Ogc-FOPDlEhl6i-Oo4ewPkb2SZRbz5g
Скрипт: script-cargo.gs
Таблиця: Logistics-Cargo
```

### Маршрути Пасажири
```
API_URL = AKfycbwd_jczf4mQRkYJgFUG2IhVGkZF-LQ9pmVEFew9gMNn2ziiKsrC3McQ4h5rfaP_FyS9
Скрипт: script-marshrut-passengers.gs
Використовується: Drivers.html (PASSENGER_API_URL, ROUTES_API_URL), Passengers.html (ROUTE_API_URL)
```

### Маршрути Посилки
```
API_URL = AKfycbwHiX1phfTEMXKQfBgHATRQ116-TFZYiJeXiswKu4eAhkjPTyRO9XMezre-LZwLPiU
Скрипт: script-marshrut-cargo.gs
Використовується: Drivers.html (DELIVERY_API_URL), Cargo.html (ROUTE_API_URL)
```

### Архів
```
API_URL = AKfycbwJLGZgYT333VdMW-nM5kPjYs2WIGGjfqkZnDJYjJxUt8nzE8GDGCPm7EzMHhcxNDOn
Скрипт: skript-arhiv.gs
```

## Аркуші Google Sheets

### Пасажири (CRM)
- `Україна-єв` — UA→EU пасажири
- `Європа-ук` — EU→UA пасажири
- `Логи` — логування дій

### Посилки (CRM)
- `Реєстрація ТТН` — UA→EU посилки
- `Виклик курєра` — EU→UA посилки
- `Логи` — логування дій

### Маршрути Посилки
- `Братислава марш.`, `Нітра марш.`, `Словаччина марш.`, `Кошице+прешов марш.`
- `Маршрути водіїв` — логи
- `Провірка розсилки` — розсилка

### Маршрути Пасажири
- Динамічні аркуші (створюються через CRM)
- `Логи`, `Провірка розсилки`

### Архів
- `Посилки`, `Посилки маршрут` — архів посилок
- `Пасажири`, `Пасажири маршрут` — архів пасажирів
- `Логи`

## Ролі

| Роль | Сторінка | Можливості |
|------|----------|------------|
| Менеджер | Passengers.html / Cargo.html | CRUD, маршрути, архівація, розсилка |
| Водій | Drivers.html | Перегляд маршрутів, оновлення статусів, навігація |

## Деплой скриптів

1. Відкрий відповідну Google Таблицю → Розширення → Apps Script
2. Заміни код на відповідний `.gs` файл
3. Deploy → New deployment → Web app
   - Execute as: Me
   - Who has access: Anyone
4. Скопіюй URL і оновити в HTML файлах

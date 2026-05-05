# milkQuality_Forms

Excel-книга с макросами (`.xlsm`) + Python-скрипт для двустороннего обмена данными с ArcGIS Feature Service (FeatureServer).

---

## Архитектура

```
milkQuality_Forms.xlsm
  └── VBA: ArcGISForms.bas   ← кнопки импорта/выгрузки на листах
  └── VBA: Module1.bas        ← отладочный запуск с перенаправлением stdout/stderr
  └── VBA: ThisWorkbook.cls   ← обработчик Worksheet_Change (колонка Dirty)

milkQuality_Forms.py          ← основная логика: запрос ArcGIS, запись в Excel, отправка правок
```

Макросы запускают Python через `Shell` / `cmd.exe /c`, передавая два аргумента:
1. `action` — строка команды (`import_f1`, `submit_f1`, `import_f2`, `submit_f2`, `import_f5`, `submit_f5`)
2. полный путь к `.xlsm`-файлу

---

## Листы и слои ArcGIS

| Лист Excel | Action-коды | FeatureServer | Описание |
|---|---|---|---|
| `Форма 1` | `import_f1` / `submit_f1` | `/2` | Акт проверки молока в танках |
| `Форма 2` | `import_f2` / `submit_f2` | `/3` | Данные молоковоза / ТТН |
| `Форма 5` | `import_f5` / `submit_f5` | `/9` | Чек-лист контроля отгрузки молока |

---

## VBA-модули

### `ArcGISForms.bas`
Основной модуль кнопок. Содержит:
- `RunPython(action)` — формирует команду и запускает `python.exe` через `Shell(..., vbNormalFocus)`.  
  Пути: `PYTHON_EXE = C:\Python311\python.exe`, скрипт ищется рядом с `.xlsm`.
- `Import_Form1/2/5` — перед импортом удаляет лист (через `DeleteSheetIfExists`), затем вызывает `RunPython`.
- `Submit_Form1/2/5` — напрямую вызывает `RunPython` с соответствующим action.
- `DeleteSheetIfExists(sheetName)` — удаляет лист без диалога подтверждения.

### `Module1.bas`
Отладочный модуль. `RunPythonWithLogs(action)` запускает скрипт через `cmd.exe /c` с явным перенаправлением `stdout` и `stderr` в отдельные `.log`-файлы (`F:\tables\<action>_stdout.log` / `_stderr.log`). Используется для диагностики, когда основное окно Shell не показывает вывод.  
`Test_submit_f5()` — ярлык для ручного запуска `submit_f5` из VBE.

> ⚠️ Пути в `Module1.bas` захардкожены на `F:\tables\` — при необходимости скорректировать.

### `ThisWorkbook.cls`
Обработчик события `Worksheet_Change`. Следит за изменениями ячеек на листах Форм 1/2/5: если изменена ячейка в строке данных (не в шапке), устанавливает `TRUE` в колонке **Dirty** этой строки. Python-скрипт при `submit_*` отправляет в ArcGIS только строки, где `Dirty = TRUE`, после чего сбрасывает флаг.

---

## Python-скрипт `milkQuality_Forms.py`

### Зависимости
```
requests
pywin32  (win32com.client)
```

### Аутентификация
Токен ArcGIS Portal получается через `generateToken`. Логин/пароль **не хранятся в файле** — берутся из переменных окружения:
```
ARCGIS_QUALITY_USER
ARCGIS_QUALITY_PASS
```

### Основные функции

| Функция | Описание |
|---|---|
| `get_token()` | Получает токен Portal через `generateToken` |
| `query_layer(url, token, ...)` | Постраничный запрос `query` к FeatureServer, возвращает список features |
| `import_sheet(wb, sheet, url, fields, sort_field)` | Импорт Форм 1/2: создаёт лист, пишет шапку и данные, форматирует даты |
| `import_f5(wb)` | Импорт Формы 5: создаёт лист с двухуровневой шапкой (группы/подгруппы), пишет 120 колонок |
| `write_sheet(ws, features, fields)` | Низкоуровневая запись данных на лист, форматирование по типу (DATE/INT/NUMBER/TEXT) |
| `push_sheet(wb, sheet, url, fields)` | Считывает строки с `Dirty=TRUE`, формирует `edits` и отправляет `updateFeatures` в ArcGIS |
| `_get_layer_oid_field(url, token)` | Запрашивает endpoint слоя и возвращает имя OID-поля |
| `log(msg)` | Пишет строку в консоль и в `.log`-файл рядом со скриптом |

### Маппинг колонок в `push_sheet` (Форма 5)
У Формы 5 множество колонок с одинаковым alias `«Оценка, балл»`, поэтому маппинг строится **по индексу колонки** (`col_to_name`), а не по alias. Для Форм 1/2 используется словарь `alias_to_name` (алиасы уникальны).

### Логирование
Все события пишутся в `milkQuality_Forms.log` рядом с `.py`-файлом.  
Формат строки: `[YYYY-MM-DD HH:MM:SS] сообщение`

---

## Настройка окружения

1. Установить Python 3.11 в `C:\Python311\`
2. Установить зависимости:
   ```bat
   C:\Python311\python.exe -m pip install requests pywin32
   ```
3. Задать переменные окружения (System → Advanced → Environment Variables):
   ```
   ARCGIS_QUALITY_USER=<логин>
   ARCGIS_QUALITY_PASS=<пароль>
   ```
4. Положить `milkQuality_Forms.py` рядом с `milkQuality_Forms.xlsm`

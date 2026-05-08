# Установка и запуск парсера MacBook (Windows)

## 1. Установить Python

Скачай и установи Python 3.12+ с https://python.org
**Важно:** при установке поставь галочку ✅ «Add Python to PATH»

Проверь:
```
python --version
```

## 2. Установить зависимости

Открой PowerShell или cmd и выполни по очереди:

```powershell
# Установка библиотек
pip install playwright pandas openpyxl playwright-stealth

# Установка браузера Chromium для Playwright
playwright install chromium
```

## 3. Запустить парсер

```powershell
python macbook_parser.py
```

Результат появится в той же папке: `macbooks_result_20260508_1430.xlsx`

## 4. Настройки (внутри скрипта)

Если хочешь видеть окно браузера, поменяй:
```python
HEADLESS = True    # → False
```

Если хочешь меньше/больше запросов:
```python
MAX_AVITO_PER_QUERY = 20   # макс. карточек с Avito
MAX_YOULA_PER_QUERY = 15   # макс. карточек с Youla
```

Если хочешь изменить поисковые запросы:
```python
SEARCH_QUERIES = [
    "MacBook M1 под ремонт",
    ...
]
```

## 5. Если Avito блокирует

Если на домашнем IP тоже банит — нужен мобильный прокси:
- **Mobileproxy.space** — 490 ₽/сутки
- **Proxy.Market** — от 170 ₽/ГБ трафика

Добавить прокси в скрипт:
```python
context = await browser.new_context(
    proxy={
        "server": "http://login:pass@ip:port",
    },
    ...
)
```

# stom PDF (HTML → PDF через WeasyPrint)

## Структура

- `data/` — входные CSV/JSON
- `templates/` — HTML-шаблоны
- `output/` — готовые PDF
- `fonts/` — (опционально) ваши шрифты `.ttf` для кириллицы
- `generate_pdf.py` — консольный генератор

## Установка зависимостей

```bash
python -m pip install -r requirements.txt
```

### Важно про Windows (WeasyPrint)
WeasyPrint на Windows может требовать системные зависимости (GTK/Pango/Cairo). Если при запуске будет ошибка импорта/рендера, поставьте зависимости по официальной инструкции WeasyPrint для Windows.

## Шрифты и кириллица

Чтобы гарантировать кириллицу на Windows/macOS, положите один из файлов в `fonts/`:

- `fonts/DejaVuSans.ttf` (рекомендуется)
- или `fonts/Roboto-Regular.ttf`

Скрипт сам попытается найти эти шрифты (сначала в `fonts/`, потом в системных папках).

## Запуск

```bash
python generate_pdf.py
```

Скрипт:

1) показывает список файлов в `data/` и шаблонов в `templates/`  
2) просит выбрать файл данных и шаблон  
3) выводит список доступных `invoice id`  
4) генерирует PDF в `output/` и открывает его системной программой

## Шаблонизация

Поддерживаются плейсхолдеры вида:

- `{{invoice_id}}`
- `{{patient_name}}`
- `{{generated_at}}`
- а также «dotted keys»: `{{patient.name}}` (если в данных вложенные объекты)

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для генерации PDF из CSV/JSON данных с использованием HTML шаблонов.
Использует WeasyPrint для конвертации HTML в PDF.
"""

import csv
import html
import json
import os
import platform
import re
import subprocess
import sys
from pathlib import Path
from string import Template

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
except ImportError:
    print("Ошибка: библиотека WeasyPrint не установлена.")
    print("Установите её командой: pip install weasyprint")
    sys.exit(1)


def get_font_config():
    """Настройка шрифтов для поддержки кириллицы."""
    font_config = FontConfiguration()
    return font_config


def get_css_styles():
    """Возвращает CSS стили для PDF с поддержкой кириллицы и таблиц."""
    css_string = """
    @page {
        size: A4;
        margin: 2cm;
        @top-right {
            content: "Страница " counter(page) " из " counter(pages);
            font-size: 10pt;
            font-family: 'DejaVu Sans', 'Roboto', 'Liberation Sans', Arial, sans-serif;
        }
    }
    
    @font-face {
        font-family: 'DejaVu Sans';
        src: local('DejaVu Sans'),
             local('DejaVuSans'),
             local('DejaVu Sans Regular');
    }
    
    @font-face {
        font-family: 'Roboto';
        src: local('Roboto'),
             local('Roboto-Regular'),
             local('Roboto Regular');
    }
    
    @font-face {
        font-family: 'Liberation Sans';
        src: local('Liberation Sans'),
             local('LiberationSans-Regular'),
             local('Liberation Sans Regular');
    }
    
    body {
        font-family: 'DejaVu Sans', 'Roboto', 'Liberation Sans', 'DejaVu Sans', Arial, 'Helvetica Neue', Helvetica, sans-serif;
        font-size: 12pt;
        line-height: 1.6;
        color: #333;
    }
    
    h1 {
        text-align: center;
        color: #2c3e50;
        margin-bottom: 30px;
        font-size: 24pt;
    }
    
    h2 {
        color: #34495e;
        margin-top: 20px;
        margin-bottom: 15px;
        font-size: 18pt;
    }
    
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px auto;
        page-break-inside: avoid;
        table-layout: fixed;
    }
    
    table th {
        background-color: #3498db;
        color: white;
        padding: 12px 8px;
        text-align: center;
        font-weight: bold;
        border: 2px solid #2980b9;
        word-wrap: break-word;
        word-break: break-word;
        overflow-wrap: break-word;
        hyphens: auto;
        font-size: 11pt;
    }
    
    table td {
        padding: 10px 8px;
        text-align: center;
        border: 1px solid #bdc3c7;
        word-wrap: break-word;
        word-break: break-word;
        overflow-wrap: break-word;
        hyphens: auto;
        font-size: 10pt;
        vertical-align: top;
    }
    
    table tbody tr:nth-child(even) {
        background-color: #f8f9fa;
    }
    
    table tbody tr:nth-child(odd) {
        background-color: #ffffff;
    }
    
    .content {
        margin: 20px 0;
    }
    
    .field-label {
        font-weight: bold;
        color: #34495e;
        margin-right: 10px;
    }
    
    .field-value {
        color: #2c3e50;
    }
    
    .invoice-info {
        background-color: #ecf0f1;
        padding: 15px;
        border-radius: 5px;
        margin: 20px 0;
    }
    """
    return CSS(string=css_string)


def scan_data_files(data_dir):
    """Сканирует директорию data и возвращает список CSV и JSON файлов."""
    data_path = Path(data_dir)
    if not data_path.exists():
        data_path.mkdir(parents=True, exist_ok=True)
        return []
    
    csv_files = list(data_path.glob("*.csv"))
    json_files = list(data_path.glob("*.json"))
    
    return {
        'csv': sorted(csv_files),
        'json': sorted(json_files)
    }


def scan_templates(templates_dir):
    """Сканирует директорию templates и возвращает список HTML шаблонов."""
    templates_path = Path(templates_dir)
    if not templates_path.exists():
        templates_path.mkdir(parents=True, exist_ok=True)
        return []
    
    html_files = list(templates_path.glob("*.html"))
    return sorted(html_files)


def read_csv_file(file_path):
    """Читает CSV файл и возвращает список словарей."""
    file_path = Path(file_path)
    
    if PANDAS_AVAILABLE:
        try:
            # Пробуем разные кодировки
            for encoding in ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251']:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    return df.to_dict('records')
                except (UnicodeDecodeError, pd.errors.EmptyDataError):
                    continue
            # Если ничего не помогло, используем utf-8-sig
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            return df.to_dict('records')
        except Exception as e:
            print(f"Ошибка при чтении CSV через pandas: {e}")
            # Fallback на стандартную библиотеку
    
    # Используем стандартную библиотеку csv
    rows = []
    try:
        for encoding in ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251']:
            try:
                with open(file_path, 'r', encoding=encoding) as csvfile:
                    sample = csvfile.read(1024)
                    csvfile.seek(0)
                    sniffer = csv.Sniffer()
                    delimiter = sniffer.sniff(sample).delimiter
                    
                    reader = csv.DictReader(csvfile, delimiter=delimiter)
                    for row in reader:
                        cleaned_row = {k.strip(): v for k, v in row.items()}
                        rows.append(cleaned_row)
                    break
            except (UnicodeDecodeError, UnicodeError):
                continue
    except Exception as e:
        print(f"Ошибка при чтении CSV: {e}")
        return []
    
    return rows


def read_json_file(file_path):
    """Читает JSON файл и возвращает данные."""
    file_path = Path(file_path)
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # Если это список, возвращаем как есть
            # Если это объект с одним ключом (например, {'invoices': [...]}), 
            # пытаемся найти список
            if isinstance(data, dict):
                # Ищем первый список в словаре
                for value in data.values():
                    if isinstance(value, list):
                        return value
                # Если списка нет, возвращаем весь словарь как список из одного элемента
                return [data]
            elif isinstance(data, list):
                return data
            else:
                return [data]
    except Exception as e:
        print(f"Ошибка при чтении JSON: {e}")
        return []


def get_invoice_ids(data):
    """Извлекает invoice id из данных (поддерживает разные варианты названий)."""
    invoice_ids = []
    
    if not data:
        return invoice_ids
    
    # Возможные названия для invoice id
    possible_keys = [
        'invoice_id', 'invoiceId', 'invoice-id', 'invoice', 
        'id', 'invoice_number', 'invoiceNumber', 'номер',
        'номер_счета', 'счет', 'check_id', 'checkId', 'check-id'
    ]
    
    for idx, item in enumerate(data):
        invoice_id = None
        
        # Если это словарь, ищем ключ
        if isinstance(item, dict):
            for key in possible_keys:
                if key.lower() in [k.lower() for k in item.keys()]:
                    # Находим точное совпадение
                    for actual_key in item.keys():
                        if actual_key.lower() == key.lower():
                            invoice_id = item[actual_key]
                            break
                    if invoice_id:
                        break
            
            # Если не нашли, используем первый ключ или индекс
            if not invoice_id:
                if item:
                    invoice_id = list(item.values())[0]
                else:
                    invoice_id = idx + 1
        else:
            invoice_id = idx + 1
        
        invoice_ids.append(str(invoice_id) if invoice_id is not None else str(idx + 1))
    
    return invoice_ids


def find_invoice_by_id(data, invoice_id):
    """Находит запись по invoice id."""
    possible_keys = [
        'invoice_id', 'invoiceId', 'invoice-id', 'invoice', 
        'id', 'invoice_number', 'invoiceNumber', 'номер',
        'номер_счета', 'счет', 'check_id', 'checkId', 'check-id'
    ]
    
    # Получаем все invoice ids для сопоставления
    invoice_ids = get_invoice_ids(data)
    
    # Ищем индекс в списке invoice_ids
    try:
        idx = invoice_ids.index(str(invoice_id))
        return data[idx]
    except ValueError:
        # Если не нашли, пытаемся найти по значению в данных
        for item in data:
            if isinstance(item, dict):
                # Проверяем все возможные ключи
                for key in possible_keys:
                    for actual_key in item.keys():
                        if actual_key.lower() == key.lower():
                            if str(item[actual_key]) == str(invoice_id):
                                return item
            
            # Если не нашли по ключу, проверяем по индексу
            if str(data.index(item) + 1) == str(invoice_id):
                return item
    
    return None


def read_template(template_path):
    """Читает HTML шаблон из файла."""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Ошибка при чтении шаблона: {e}")
        return None


def substitute_template(template_str, data):
    """Подставляет данные в шаблон используя Template."""
    template = Template(template_str)
    
    # Подготавливаем данные для подстановки
    safe_data = {}
    for key, value in data.items():
        # Escape $ символы, если они есть в данных
        if value and '$' in str(value):
            safe_value = str(value).replace('$', '$$')
            safe_data[key] = safe_value
        else:
            safe_data[key] = value if value else ''
    
    try:
        return template.safe_substitute(safe_data)
    except Exception as e:
        print(f"Ошибка при подстановке данных в шаблон: {e}")
        raise


def create_html_content(data, template_str):
    """Создает HTML контент из данных и шаблона."""
    # Создаем таблицу со всеми данными
    table_html = "<table>\n<thead>\n<tr>\n"
    
    # Заголовки таблицы
    for key in data.keys():
        safe_key = html.escape(str(key))
        table_html += f"    <th>{safe_key}</th>\n"
    
    table_html += "</tr>\n</thead>\n<tbody>\n<tr>\n"
    
    # Значения таблицы
    for value in data.values():
        safe_value = html.escape(str(value)) if value else ''
        table_html += f"    <td>{safe_value}</td>\n"
    
    table_html += "</tr>\n</tbody>\n</table>\n"
    
    # Добавляем таблицу в данные для подстановки в шаблон
    template_data = data.copy()
    template_data['table_content'] = table_html
    
    # Подставляем все данные в шаблон
    html_content = substitute_template(template_str, template_data)
    
    return html_content


def generate_pdf(html_content, output_path, css_styles, font_config):
    """Генерирует PDF файл из HTML контента."""
    try:
        html = HTML(string=html_content, encoding='utf-8')
        html.write_pdf(
            output_path,
            stylesheets=[css_styles],
            font_config=font_config
        )
        return True
    except Exception as e:
        # На некоторых консолях Windows символ ✗ может вызывать UnicodeEncodeError,
        # поэтому выводим сообщение без специальных символов.
        print(f"Ошибка при создании PDF: {e}")
        return False


def open_pdf(file_path):
    """Открывает PDF файл в системной программе просмотра."""
    try:
        system = platform.system()
        file_path_abs = str(Path(file_path).absolute())
        
        if system == 'Windows':
            os.startfile(file_path_abs)
        elif system == 'Darwin':  # macOS
            subprocess.run(['open', file_path_abs], check=False)
        elif system == 'Linux':
            subprocess.run(['xdg-open', file_path_abs], check=False)
        else:
            print(f"Неизвестная ОС: {system}. Откройте файл вручную: {file_path_abs}")
    except Exception as e:
        print(f"Предупреждение: не удалось автоматически открыть PDF: {e}")
        print(f"Откройте файл вручную: {Path(file_path).absolute()}")


def display_menu(title, items):
    """Выводит меню с нумерацией."""
    print(f"\n{'='*60}")
    print(f"  {title}")
    print(f"{'='*60}")
    for idx, item in enumerate(items, start=1):
        print(f"  {idx}. {item}")
    print(f"{'='*60}\n")


def get_user_choice(max_choice, prompt="Выберите вариант"):
    """Получает выбор пользователя с проверкой."""
    while True:
        try:
            choice = input(f"{prompt} (1-{max_choice}): ").strip()
            choice_num = int(choice)
            if 1 <= choice_num <= max_choice:
                return choice_num - 1  # Возвращаем индекс (0-based)
            else:
                print(f"Пожалуйста, введите число от 1 до {max_choice}")
        except ValueError:
            print("Пожалуйста, введите число")
        except KeyboardInterrupt:
            print("\n\nПрервано пользователем.")
            sys.exit(0)


def main():
    """Основная функция."""
    # Настройки
    data_dir = Path('data')
    templates_dir = Path('templates')
    output_dir = Path('output')
    
    # Создаем необходимые директории
    data_dir.mkdir(exist_ok=True)
    templates_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)
    
    print("\n" + "="*60)
    print("  ГЕНЕРАТОР PDF ИЗ CSV/JSON ДАННЫХ")
    print("="*60)
    
    # Сканируем файлы
    print("\nСканирую файлы...")
    data_files = scan_data_files(data_dir)
    templates = scan_templates(templates_dir)
    
    all_data_files = data_files['csv'] + data_files['json']
    
    if not all_data_files:
        print("\n⚠ Ошибка: в директории 'data' не найдено CSV или JSON файлов!")
        print("Создайте файлы с данными в директории 'data/'")
        return
    
    if not templates:
        print("\n⚠ Ошибка: в директории 'templates' не найдено HTML шаблонов!")
        print("Создайте HTML шаблоны в директории 'templates/'")
        return
    
    # Выводим доступные файлы данных
    data_file_names = [f"{f.name} ({'CSV' if f in data_files['csv'] else 'JSON'})" 
                       for f in all_data_files]
    display_menu("ДОСТУПНЫЕ ФАЙЛЫ С ДАННЫМИ", data_file_names)
    
    # Пользователь выбирает файл данных
    data_choice = get_user_choice(len(all_data_files), "Выберите файл с данными")
    selected_data_file = all_data_files[data_choice]
    
    print(f"\n✓ Выбран файл: {selected_data_file.name}")
    
    # Читаем данные
    print("Читаю данные...")
    if selected_data_file.suffix.lower() == '.csv':
        data = read_csv_file(selected_data_file)
    else:
        data = read_json_file(selected_data_file)
    
    if not data:
        print("⚠ Ошибка: не удалось прочитать данные или файл пуст!")
        return
    
    print(f"✓ Загружено записей: {len(data)}")
    
    # Выводим доступные шаблоны
    template_names = [t.name for t in templates]
    display_menu("ДОСТУПНЫЕ HTML ШАБЛОНЫ", template_names)
    
    # Пользователь выбирает шаблон
    template_choice = get_user_choice(len(templates), "Выберите HTML шаблон")
    selected_template = templates[template_choice]
    
    print(f"\n✓ Выбран шаблон: {selected_template.name}")
    
    # Читаем шаблон
    template_str = read_template(selected_template)
    if not template_str:
        print("⚠ Ошибка: не удалось прочитать шаблон!")
        return
    
    # Извлекаем invoice ids
    invoice_ids = get_invoice_ids(data)
    
    # Выводим список invoice ids
    invoice_display = [f"Invoice ID: {invoice_id}" for invoice_id in invoice_ids]
    display_menu(f"ДОСТУПНЫЕ СЧЕТА (INVOICE ID) - Всего: {len(invoice_ids)}", invoice_display)
    
    # Пользователь выбирает invoice id
    invoice_choice = get_user_choice(len(invoice_ids), "Выберите invoice ID для генерации PDF")
    selected_invoice_id = invoice_ids[invoice_choice]
    
    print(f"\n✓ Выбран Invoice ID: {selected_invoice_id}")
    
    # Находим данные для выбранного invoice
    invoice_data = find_invoice_by_id(data, selected_invoice_id)
    
    if not invoice_data:
        print("⚠ Ошибка: не удалось найти данные для выбранного invoice ID!")
        return
    
    # Преобразуем данные в словарь, если нужно
    if not isinstance(invoice_data, dict):
        invoice_data = {'data': invoice_data}
    
    # Создаем HTML контент
    print("\nГенерирую HTML...")
    html_content = create_html_content(invoice_data, template_str)
    
    # Определяем имя выходного файла
    safe_filename = re.sub(r'[<>:"/\\|?*]', '', str(selected_invoice_id))
    safe_filename = safe_filename.replace(' ', '_')
    pdf_path = output_dir / f"invoice_{safe_filename}.pdf"
    
    # Настраиваем стили и шрифты
    css_styles = get_css_styles()
    font_config = get_font_config()
    
    # Генерируем PDF
    print("Генерирую PDF...")
    if generate_pdf(html_content, str(pdf_path), css_styles, font_config):
        print(f"✓ PDF создан: {pdf_path}")
        
        # Открываем PDF
        print("\nОткрываю PDF...")
        open_pdf(pdf_path)
        print("\nГотово!")
    else:
        print("Ошибка при создании PDF!")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nПрервано пользователем.")
        sys.exit(0)
from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from html import escape
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from weasyprint import CSS, HTML


ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "data"
TEMPLATES_DIR = ROOT_DIR / "templates"
OUTPUT_DIR = ROOT_DIR / "output"
FONTS_DIR = ROOT_DIR / "fonts"


SUPPORTED_DATA_EXTS = {".csv", ".json"}


def _configure_windows_utf8_console() -> None:
    """
    Делает вывод/ввод кириллицы более стабильным в Windows-консоли.
    Не меняет системные настройки, только кодировки потоков Python.
    """
    if not sys.platform.startswith("win"):
        return
    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    try:
        sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass
    try:
        sys.stdin.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass


def _ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    FONTS_DIR.mkdir(parents=True, exist_ok=True)


def _print_header(title: str) -> None:
    print("\n" + "=" * 72)
    print(title)
    print("=" * 72)


def _choose_from_list(prompt: str, items: Sequence[str]) -> str:
    if not items:
        raise ValueError("Список пуст")

    while True:
        print()
        for i, item in enumerate(items, start=1):
            print(f"{i:>2}. {item}")
        raw = input(f"\n{prompt} (1-{len(items)}): ").strip()
        if not raw.isdigit():
            print("Введите номер варианта.")
            continue
        idx = int(raw)
        if not (1 <= idx <= len(items)):
            print("Неверный номер.")
            continue
        return items[idx - 1]


def _list_files_sorted(dir_path: Path, exts: set[str]) -> List[Path]:
    if not dir_path.exists():
        return []
    files: List[Path] = []
    for p in dir_path.iterdir():
        if p.is_file() and p.suffix.lower() in exts:
            files.append(p)
    return sorted(files, key=lambda x: x.name.lower())


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _try_import_pandas():
    try:
        import pandas as pd  # type: ignore

        return pd
    except Exception:
        return None


def _load_csv_records(path: Path) -> List[Dict[str, Any]]:
    """
    Returns list of dict rows. Prefers pandas (if installed), otherwise csv.DictReader.
    """
    pd = _try_import_pandas()
    if pd is not None:
        df = pd.read_csv(path)
        # pandas can produce NaN; normalize to None / str later during rendering
        return df.to_dict(orient="records")

    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        return [dict(row) for row in reader]


def _load_json_records(path: Path) -> List[Dict[str, Any]]:
    """
    Supports:
    - list of objects: [{...}, {...}]
    - object with 'invoices' list: {"invoices": [{...}]}
    - dict keyed by id: {"INV-1": {...}, "INV-2": {...}}
    """
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]

    if isinstance(data, dict):
        if "invoices" in data and isinstance(data["invoices"], list):
            return [x for x in data["invoices"] if isinstance(x, dict)]

        # dict-of-dicts keyed by id
        values: List[Dict[str, Any]] = []
        for k, v in data.items():
            if isinstance(v, dict):
                row = dict(v)
                if "invoice_id" not in row and "invoiceId" not in row and "id" not in row:
                    row["invoice_id"] = k
                values.append(row)
        if values:
            return values

    raise ValueError(
        "JSON формат не распознан. Ожидается список объектов или объект с ключом 'invoices'."
    )


def _load_records(path: Path) -> List[Dict[str, Any]]:
    ext = path.suffix.lower()
    if ext == ".csv":
        return _load_csv_records(path)
    if ext == ".json":
        return _load_json_records(path)
    raise ValueError(f"Неподдерживаемый тип данных: {ext}")


INVOICE_ID_CANDIDATES = (
    "invoice_id",
    "invoiceId",
    "invoice id",
    "invoice",
    "InvoiceId",
    "InvoiceID",
    "id",
    "Id",
    "ID",
)


def _extract_invoice_id(record: Dict[str, Any]) -> Optional[str]:
    for k in INVOICE_ID_CANDIDATES:
        if k in record and record[k] is not None:
            val = record[k]
            s = str(val).strip()
            if s and s.lower() != "nan":
                return s
    return None


def _records_by_invoice_id(records: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    auto_i = 1
    for r in records:
        if not isinstance(r, dict):
            continue
        inv = _extract_invoice_id(r)
        if not inv:
            inv = f"AUTO-{auto_i}"
            auto_i += 1
        # last wins
        out[inv] = r
    return out


_TPL_VAR_RE = re.compile(r"\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}")


def _get_nested(data: Dict[str, Any], dotted_key: str) -> Any:
    cur: Any = data
    for part in dotted_key.split("."):
        if isinstance(cur, dict) and part in cur:
            cur = cur[part]
        else:
            return None
    return cur


def render_html_template(template_html: str, context: Dict[str, Any]) -> str:
    """
    Minimal templating:
    - Replaces {{key}} with HTML-escaped value
    - Supports dotted keys: {{patient.name}}
    """

    def repl(m: re.Match) -> str:
        key = m.group(1)
        val = _get_nested(context, key)
        if val is None:
            return ""
        # normalize NaN from pandas
        s = str(val)
        if s.lower() == "nan":
            return ""
        return escape(s)

    return _TPL_VAR_RE.sub(repl, template_html)


def _find_font_file() -> Optional[Path]:
    """
    Preference:
    - project fonts/
    - common system locations (Windows/macOS)
    """
    candidates = [
        FONTS_DIR / "DejaVuSans.ttf",
        FONTS_DIR / "DejaVuSansCondensed.ttf",
        FONTS_DIR / "Roboto-Regular.ttf",
        FONTS_DIR / "Roboto.ttf",
    ]

    # Windows
    win = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    candidates += [
        win / "DejaVuSans.ttf",
        win / "DejaVuSansCondensed.ttf",
        win / "Roboto-Regular.ttf",
        win / "Roboto.ttf",
        win / "arial.ttf",
        win / "segoeui.ttf",
    ]

    # macOS
    candidates += [
        Path("/Library/Fonts/DejaVu Sans.ttf"),
        Path("/Library/Fonts/DejaVuSans.ttf"),
        Path("/Library/Fonts/Roboto-Regular.ttf"),
        Path("/System/Library/Fonts/Supplemental/Roboto-Regular.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf"),
    ]

    for p in candidates:
        try:
            if p.exists() and p.is_file():
                return p
        except Exception:
            continue
    return None


def _build_css(font_path: Optional[Path]) -> CSS:
    # WeasyPrint supports @font-face with file:// URLs
    font_face = ""
    font_family = "DejaVuSans"
    if font_path is not None:
        font_url = font_path.resolve().as_uri()
        font_face = f"""
@font-face {{
  font-family: "{font_family}";
  src: url("{font_url}");
}}
"""
    else:
        # fallback to system fonts; still keep Cyrillic-friendly candidates
        font_family = 'DejaVu Sans, Roboto, "Segoe UI", Arial, sans-serif'

    base = f"""
{font_face}
html, body {{
  font-family: {font_family};
  font-size: 12pt;
  color: #111;
}}
@page {{
  size: A4;
  margin: 12mm;
}}
"""
    return CSS(string=base)


def _safe_filename(s: str) -> str:
    # keep ASCII, digits, dash/underscore; replace others
    out = re.sub(r"[^a-zA-Z0-9._-]+", "_", s.strip())
    return out or "invoice"


def _open_file_in_system_viewer(path: Path) -> None:
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
            return
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
            return
        # fallback for other OS
        subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        print(f"Не удалось автоматически открыть PDF: {e}")


def main() -> int:
    _configure_windows_utf8_console()
    _ensure_dirs()

    _print_header("Поиск данных и шаблонов")
    data_files = _list_files_sorted(DATA_DIR, SUPPORTED_DATA_EXTS)
    template_files = _list_files_sorted(TEMPLATES_DIR, {".html", ".htm"})

    print(f"Директория данных:     {DATA_DIR}")
    print(f"Директория шаблонов:   {TEMPLATES_DIR}")
    print(f"Директория вывода PDF: {OUTPUT_DIR}")

    print("\nДоступные файлы данных:")
    if not data_files:
        print("  (нет файлов .csv/.json в папке data)")
    else:
        for i, p in enumerate(data_files, start=1):
            print(f"  {i:>2}. {p.name}")

    print("\nДоступные HTML-шаблоны:")
    if not template_files:
        print("  (нет файлов .html/.htm в папке templates)")
    else:
        for i, p in enumerate(template_files, start=1):
            print(f"  {i:>2}. {p.name}")

    if not data_files or not template_files:
        print(
            "\nДобавьте хотя бы один файл данных в `data/` и один HTML-шаблон в `templates/`, затем запустите снова."
        )
        return 2

    chosen_data = _choose_from_list(
        "Выберите файл данных", [p.name for p in data_files]
    )
    chosen_tpl = _choose_from_list(
        "Выберите HTML-шаблон", [p.name for p in template_files]
    )

    data_path = DATA_DIR / chosen_data
    tpl_path = TEMPLATES_DIR / chosen_tpl

    _print_header("Загрузка данных")
    print(f"Файл данных:   {data_path.name}")
    print(f"HTML-шаблон:   {tpl_path.name}")

    try:
        records = _load_records(data_path)
    except Exception as e:
        print(f"Ошибка чтения данных: {e}")
        return 3

    by_id = _records_by_invoice_id(records)
    invoice_ids = sorted(by_id.keys(), key=lambda x: x.lower())

    _print_header("Доступные чеки (invoice id)")
    if not invoice_ids:
        print("Не найдено ни одной записи.")
        return 4

    chosen_id = _choose_from_list("Выберите invoice id", invoice_ids)
    record = by_id[chosen_id]

    _print_header("Генерация PDF")
    template_html = _read_text(tpl_path)
    # enrich context with a couple of useful fields
    context = dict(record)
    context.setdefault("invoice_id", chosen_id)
    context.setdefault("generated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    final_html = render_html_template(template_html, context)

    font_path = _find_font_file()
    if font_path is None:
        print(
            "Шрифт DejaVuSans/Roboto не найден. Для надежной кириллицы положите TTF в `fonts/` "
            "(например, `fonts/DejaVuSans.ttf`)."
        )
    else:
        print(f"Используем шрифт: {font_path}")

    css = _build_css(font_path)
    out_name = f"{_safe_filename(chosen_id)}.pdf"
    out_path = OUTPUT_DIR / out_name

    try:
        HTML(string=final_html, base_url=str(TEMPLATES_DIR)).write_pdf(
            str(out_path), stylesheets=[css]
        )
    except Exception as e:
        print(f"Ошибка генерации PDF (WeasyPrint): {e}")
        return 5

    print(f"PDF сохранен: {out_path}")
    _open_file_in_system_viewer(out_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

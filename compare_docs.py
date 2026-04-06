"""
Скрипт для сравнения документов из двух папок:
  - folder_tz/   — документы из ТЗ 
  - folder_lk/   — документы из ЛК Клиента

Поддерживаемые форматы: .docx, .doc
"""

import os
import re
import subprocess
import difflib
from pathlib import Path
from dataclasses import dataclass, field

try:
    from docx import Document
except ImportError:
    raise SystemExit("Установите зависимость: pip install python-docx")

try:
    import pdfplumber
except ImportError:
    raise SystemExit("Установите зависимость: pip install pdfplumber")


# ── Настройки ──────────────────────────────────────────────────────────────────

FOLDER_TZ = Path("folder_tz")   # папка с документами из ТЗ
FOLDER_LK = Path("folder_lk")   # папка с документами из ЛК Клиента

EXTENSIONS = {".docx", ".doc", ".pdf"}

DEBUG = False


# ── Извлечение текста ──────────────────────────────────────────────────────────

def extract_text_docx(path: Path) -> str:

    from docx.text.paragraph import Paragraph as DocxParagraph
    from docx.table import Table as DocxTable

    doc = Document(path)
    parts = []

    for element in doc.element.body:
        tag = element.tag.split("}")[-1]  #
        if tag == "p":
            
            text = DocxParagraph(element, doc).text.replace("\n", " ").strip()
            if text:
                parts.append(text)
        elif tag == "tbl":
            table = DocxTable(element, doc)
            for row in table.rows:

                seen_in_row: set = set()
                cell_texts = []
                for cell in row.cells:
                    if cell._tc not in seen_in_row:
                        seen_in_row.add(cell._tc)
                        t = cell.text.replace("\n", " ").strip()
                        if t:
                            cell_texts.append(t)
                if cell_texts:
                    parts.append(" ".join(cell_texts))

    return "\n".join(parts)


def extract_text_doc_antiword(path: Path) -> str:

    result = subprocess.run(
        ["antiword", str(path)],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.returncode != 0:
        raise RuntimeError(f"antiword error: {result.stderr.strip()}")
    return result.stdout



def extract_text_pdf(path: Path) -> str:

    with pdfplumber.open(path) as pdf:
        pages = [page.extract_text() or "" for page in pdf.pages]

    text = "\n".join(pages)
    return re.sub(r"-\s*\n\s*", "", text)


def extract_text(path: Path) -> str:

    ext = path.suffix.lower()
    if ext == ".docx":
        return extract_text_docx(path)
    elif ext == ".doc":

            return extract_text_doc_antiword(path)
    elif ext == ".pdf":
        return extract_text_pdf(path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}")


# ── Сравнение ──────────────────────────────────────────────────────────────────

@dataclass
class FileResult:
    name: str
    status: str          # "identical" | "different" | "only_in_tz" | "only_in_lk" | "error"
    diff_lines: list[str] = field(default_factory=list)
    error: str = ""


def normalize_text(text: str) -> str:

    text = re.sub(r"-\s*\n\s*", "", text)
    text = re.sub(r"[^\w]", " ", text, flags=re.UNICODE)
    return re.sub(r" +", " ", text).strip()


def compare_texts(text_tz: str, text_lk: str, filename: str) -> list[str]:

    norm_tz = normalize_text(text_tz)
    norm_lk = normalize_text(text_lk)

    words_tz = re.findall(r'\w+', norm_tz, flags=re.UNICODE)
    words_lk = re.findall(r'\w+', norm_lk, flags=re.UNICODE)

    diff = list(difflib.unified_diff(
        words_tz, words_lk,
        fromfile=f"ТЗ/{filename}",
        tofile=f"ЛК/{filename}",
        lineterm="",
        n=2, 
    ))
    return diff


def dump_debug(name: str, text_tz: str, text_lk: str) -> None:

    debug_dir = Path("debug")
    debug_dir.mkdir(exist_ok=True)
    stem = Path(name).stem
    (debug_dir / f"{stem}_tz.txt").write_text(text_tz, encoding="utf-8")
    (debug_dir / f"{stem}_lk.txt").write_text(text_lk, encoding="utf-8")


def compare_file(name: str, path_tz: Path, path_lk: Path) -> FileResult:
    try:
        text_tz = extract_text(path_tz)
        text_lk = extract_text(path_lk)
    except Exception as e:
        return FileResult(name=name, status="error", error=str(e))

    if DEBUG:
        dump_debug(name, normalize_text(text_tz), normalize_text(text_lk))

    diff = compare_texts(text_tz, text_lk, name)
    if diff:
        return FileResult(name=name, status="different", diff_lines=diff)
    return FileResult(name=name, status="identical")


# ── Основная логика ────────────────────────────────────────────────────────────

def collect_files(folder: Path) -> dict[str, Path]:

    return {
        f.stem: f
        for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSIONS
    }


def run_comparison() -> list[FileResult]:
    results = []

    files_tz = collect_files(FOLDER_TZ)
    files_lk = collect_files(FOLDER_LK)

    all_names = sorted(set(files_tz) | set(files_lk))

    for name in all_names:
        in_tz = name in files_tz
        in_lk = name in files_lk

        if in_tz and in_lk:
            results.append(compare_file(name, files_tz[name], files_lk[name]))
        elif in_tz:
            results.append(FileResult(name=name, status="only_in_tz"))
        else:
            results.append(FileResult(name=name, status="only_in_lk"))

    return results


# ── Вывод отчёта ───────────────────────────────────────────────────────────────

STATUS_LABELS = {
    "identical":  "✓ Совпадает с эталоном",
    "different":  "✗ ОТЛИЧАЕТСЯ от эталона",
    "only_in_tz": "✗ ОТСУТСТВУЕТ в ЛК (есть в эталоне ТЗ)",
    "only_in_lk": "! Лишний файл в ЛК (нет в эталоне ТЗ)",
    "error":      "? Ошибка чтения",
}


def build_report_lines(results: list[FileResult]) -> tuple[list[str], bool]:

    identical = sum(1 for r in results if r.status == "identical")
    different = sum(1 for r in results if r.status == "different")
    only_tz   = sum(1 for r in results if r.status == "only_in_tz")
    only_lk   = sum(1 for r in results if r.status == "only_in_lk")
    errors    = sum(1 for r in results if r.status == "error")
    has_issues = different > 0 or only_tz > 0 or only_lk > 0 or errors > 0

    lines = []
    lines.append("=" * 60)
    lines.append("  ПРОВЕРКА ДОКУМЕНТОВ ЛК ПО ЭТАЛОНУ ТЗ")
    lines.append("=" * 60)
    lines.append(f"  Всего файлов в (ТЗ): {len([r for r in results if r.status != 'only_in_lk'])}")
    lines.append(f"  Совпадают с ТЗ:        {identical}")
    lines.append(f"  Отличаются от ТЗ:       {different}")
    lines.append(f"  Отсутствуют в ЛК:            {only_tz}")
    lines.append(f"  Лишние файлы в ЛК:           {only_lk}")
    lines.append(f"  Ошибки чтения:               {errors}")
    lines.append("=" * 60)


    sorted_results = sorted(results, key=lambda r: (r.status == "identical", r.name))

    for r in sorted_results:
        label = STATUS_LABELS[r.status]
        lines.append(f"\n[{label}] {r.name}")

        if r.status == "different":
            lines.append("  Отклонения ЛК от эталона ТЗ:")
            lines.append("  (- удалено из ТЗ  |  + добавлено в ЛК)")
            for line in r.diff_lines:

                if line.startswith("---") or line.startswith("+++") or line.startswith("@@"):
                    continue

                if line.startswith("+") or line.startswith("-"):
                    lines.append("    " + line)

        elif r.status == "error":
            lines.append(f"  Ошибка: {r.error}")

    lines.append("\n" + "=" * 60)
    if not has_issues:
        lines.append("  Все документы ЛК соответствуют эталону ТЗ.")
    else:
        lines.append("  ВНИМАНИЕ: Документы ЛК не соответствуют эталону ТЗ!")
    lines.append("=" * 60)

    return lines, has_issues


def print_report(results: list[FileResult]) -> bool:
    """Выводит отчёт в консоль. Возвращает True, если есть расхождения."""
    lines, has_issues = build_report_lines(results)
    print("\n".join(lines))
    return has_issues


def save_report(results: list[FileResult], output_path: Path) -> None:
    """Сохраняет отчёт о расхождениях в текстовый файл."""
    lines, _ = build_report_lines(results)
    output_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nОтчёт сохранён: {output_path.resolve()}")




REPORT_FILE = Path("report.txt")  

if __name__ == "__main__":
    for folder in (FOLDER_TZ, FOLDER_LK):
        if not folder.exists():
            raise SystemExit(f"Папка не найдена: {folder.resolve()}")

    results = run_comparison()
    has_issues = print_report(results)

    if has_issues:
        save_report(results, REPORT_FILE)

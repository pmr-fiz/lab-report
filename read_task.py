"""
read_task.py — извлекает текст задания лабы из .docx или .pdf.

Лабы приходят как Word или PDF (в markdown их не дают). Скрипт вытаскивает
из них текст, который дальше читает агент. Никаких картинок/OCR —
только текст, легко и быстро.

Использование:
    python read_task.py labs/1/task.docx
    python read_task.py labs/1/task.pdf

Зависимости: python-docx, pypdf.
"""

import os
import sys


def read_docx(path):
    from docx import Document
    doc = Document(path)
    lines = []
    for p in doc.paragraphs:
        if p.text.strip():
            lines.append(p.text)
    for t in doc.tables:
        lines.append("\n[таблица]")
        for row in t.rows:
            lines.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(lines)


def read_pdf(path):
    from pypdf import PdfReader
    reader = PdfReader(path)
    text = "\n".join((pg.extract_text() or "") for pg in reader.pages).strip()
    if len(text) < 40:
        return ("[Не удалось извлечь текст — вероятно, PDF из сканов/картинок. "
                "Дай задание в .docx или в PDF с текстовым слоем.]")
    return text


def read_task(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return read_docx(path)
    if ext == ".pdf":
        return read_pdf(path)
    if ext in (".md", ".txt"):
        with open(path, encoding="utf-8") as f:
            return f.read()
    raise SystemExit(f"Поддерживаются только .docx и .pdf (получено: {ext})")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Использование: python read_task.py <task.docx|task.pdf>")
    print(read_task(sys.argv[1]))

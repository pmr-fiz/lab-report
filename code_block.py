"""
Переиспользуемый компонент для агента-оформителя отчётов.

add_code_block(doc, code, language)  -> вставляет код настоящим текстом,
    подсвеченным по синтаксису, на тёмном фоне (цвета темы VS Code Dark+).
add_output_block(doc, text, caption) -> вставляет вывод программы отдельным
    тёмным блоком (без подсветки, просто моноширинный текст).

Код получается ВЫДЕЛЯЕМЫМ и КОПИРУЕМЫМ (в отличие от скриншота).
"""

from docx.shared import Pt, RGBColor, Twips
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from pygments import lex
from pygments.lexers import get_lexer_by_name
from pygments.token import Token

# --- Палитра под тему VS Code Dark+ (близко к твоим скриншотам) ---
BG        = "1E1E1E"   # фон блока
DEFAULT   = "D4D4D4"   # обычный текст / операторы
KEYWORD   = "569CD6"   # def, for, if, else, return, in, import ...
FUNC      = "DCDCAA"   # имена функций и вызовы (print, enumerate, float)
STRING    = "CE9178"   # строки
NUMBER    = "B5CEA8"   # числа
COMMENT   = "6A9955"   # комментарии
NAME      = "9CDCFE"   # обычные переменные
OUTPUT_FG = "D4D4D4"   # цвет текста вывода

FONT = "Consolas"      # моноширинный; запасной — Courier New
SIZE = 9               # кегль кода


def _token_color(tok):
    """Сопоставляем тип токена Pygments -> hex-цвет темы."""
    if tok in Token.Comment:
        return COMMENT
    if tok in Token.String:
        return STRING
    if tok in Token.Number:
        return NUMBER
    if tok in Token.Keyword:
        return KEYWORD
    if tok in Token.Name.Function or tok in Token.Name.Builtin:
        return FUNC
    if tok in Token.Name:
        return NAME
    if tok in Token.Operator or tok in Token.Punctuation:
        return DEFAULT
    return DEFAULT


def _shade(element, fill):
    """Заливка фоном (ячейки таблицы) через прямой XML."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    element.append(shd)


def _no_space(paragraph):
    """Убираем интервалы между строками кода, делаем плотную вёрстку."""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _dark_cell(doc, width_twips=9200):
    """Одноячеечная таблица с тёмной заливкой — контейнер под блок."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Twips(width_twips)
    # ширина таблицы + ячейки (докс требует и то, и то)
    tblPr = table._tbl.tblPr
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(width_twips))
    tblPr.append(tblW)
    _shade(cell._tc.get_or_add_tcPr(), BG)
    # внутренние поля ячейки, чтобы текст не липнул к краю
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", 80), ("bottom", 80), ("left", 120), ("right", 120)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:type"), "dxa")
        m.set(qn("w:w"), str(val))
        tcMar.append(m)
    cell._tc.get_or_add_tcPr().append(tcMar)
    # убираем дефолтный пустой абзац в ячейке
    cell.paragraphs[0]._p.getparent().remove(cell.paragraphs[0]._p)
    return cell


def _add_run(paragraph, text, hexcolor):
    run = paragraph.add_run(text)
    run.font.name = FONT
    run.font.size = Pt(SIZE)
    run.font.color.rgb = RGBColor.from_string(hexcolor)
    # закрепляем моноширинный шрифт и для кириллицы
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), FONT)
    return run


def add_code_block(doc, code, language="python"):
    """Вставляет подсвеченный код настоящим текстом на тёмном фоне."""
    cell = _dark_cell(doc)
    lexer = get_lexer_by_name(language)

    # Первый абзац
    para = cell.add_paragraph()
    _no_space(para)

    for tok, value in lex(code, lexer):
        color = _token_color(tok)
        # value может содержать переносы строк -> делим на отдельные абзацы
        parts = value.split("\n")
        for idx, part in enumerate(parts):
            if idx > 0:
                para = cell.add_paragraph()
                _no_space(para)
            if part:
                _add_run(para, part, color)
    return cell


def add_output_block(doc, text, caption="Результат работы программы:"):
    """Вставляет вывод программы отдельным тёмным блоком (без подсветки)."""
    if caption:
        cap = doc.add_paragraph()
        cap.paragraph_format.space_before = Pt(6)
        cap.paragraph_format.space_after = Pt(2)
        r = cap.add_run(caption)
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)

    cell = _dark_cell(doc)
    para = cell.add_paragraph()
    _no_space(para)
    for idx, line in enumerate(text.split("\n")):
        if idx > 0:
            para = cell.add_paragraph()
            _no_space(para)
        if line:
            _add_run(para, line, OUTPUT_FG)
    return cell

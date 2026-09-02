"""
builder.py
Собирает финальный .docx для заказчика: акцентная плашка заголовка,
выделенные разделы с цветной полосой, читаемые таблицы без грубой сетки,
блоки-выноски для "Важно"/пробелов, колонтитул с номером страницы.
"""
from __future__ import annotations

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# --- палитра (перекликается с веб-версией инструмента) ------------------
ACCENT = RGBColor(0x1D, 0x6E, 0x28)        # глубокий зелёный — заголовки, акценты
ACCENT_HEX = "1D6E28"
RULE_HEX = "2FA83D"                        # яркий зелёный — акцентные полосы
TITLE_BG_HEX = "EAF8EC"                    # светло-зелёный фон плашки заголовка
TABLE_HEADER_HEX = "14261A"                # тёмный графитово-зелёный — шапка таблиц
MUTED = RGBColor(0x5B, 0x64, 0x72)
MUTED_HEX = "5B6472"
BORDER_HEX = "DCE3E9"                      # светло-серый — границы таблиц
ALT_ROW_HEX = "F3F9F4"                     # чередующаяся заливка строк (зеленоватая)
CALLOUT_BG_HEX = "FFF6E8"                  # тёплый фон для "Важно"
CALLOUT_BORDER_HEX = "E0A93B"
WARN = RGBColor(0xA6, 0x2C, 0x2C)
WARN_BG_HEX = "FBEAEA"
WARN_BORDER_HEX = "C0504D"
GROOMING_BG_HEX = "EAF1FB"                 # синий — для "нужен грумминг" (отличается от акцентного зелёного)
GROOMING_BORDER_HEX = "2E6FBE"
GROOMING_TEXT_COLOR = RGBColor(0x1D, 0x4E, 0x8C)
BUSINESS_CASE_BG_HEX = "F1EFFC"            # мягкий фиолетовый — для раздела "Бизнес-кейс"
BUSINESS_CASE_BORDER_HEX = "8B7FD6"
BUSINESS_CASE_TEXT_COLOR = RGBColor(0x4A, 0x3F, 0x91)
BUSINESS_CASE_HEADINGS = ("бизнес-кейс", "бизнес кейс")

HEADING_FONT = "Calibri"                   # без засечек, как и в веб-версии — единая гарнитура
BODY_FONT = "Calibri"

CALLOUT_TRIGGER_RE = re.compile(r"^[^a-zA-Zа-яА-ЯёЁ]*(важно|внимание)[:!]?\s*$", re.IGNORECASE)
GROOMING_TRIGGER_RE = re.compile(r"грум{1,2}инг", re.IGNORECASE)
_CELL_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|(?=\()")


def _split_cell_grooming_runs(text: str) -> list[tuple[str, bool]]:
    """Для ячейки таблицы: режет текст на куски (предложения) и помечает,
    какие из них — вопрос на грумминг. В отличие от extract_grooming_blocks
    (tracker_parser.py) текст из ячейки никуда не выносится — таблица не
    резиновая, вырезать середину ячейки значит оставить дыру в данных.
    Здесь только точечно подсвечивается нужный кусок прямо на месте.

    Вопрос на грумминг часто пишут в скобках одним заходом, но с несколькими
    предложениями внутри — только первое из них обычно содержит само слово
    "грумминг" ("(Вопрос на грумминг: ...? только если...?)"). Чтобы не
    подсвечивать вопрос наполовину, отслеживаем баланс скобок: если
    грумминг-предложение открывает скобку и не закрывает её сразу — все
    следующие предложения тоже считаются частью того же вопроса, пока
    скобка не закроется."""
    if not GROOMING_TRIGGER_RE.search(text):
        return [(text, False)]
    sentences = [s for s in _CELL_SENTENCE_SPLIT_RE.split(text) if s.strip()]
    if len(sentences) <= 1:
        return [(text, True)]

    runs: list[tuple[str, bool]] = []
    paren_depth = 0
    inside_grooming_parens = False
    for s in sentences:
        is_g = bool(GROOMING_TRIGGER_RE.search(s)) or inside_grooming_parens
        opens, closes = s.count("("), s.count(")")
        if is_g and not inside_grooming_parens and opens > closes:
            inside_grooming_parens = True
        paren_depth = max(0, paren_depth + opens - closes)
        if inside_grooming_parens and paren_depth == 0:
            inside_grooming_parens = False
        if runs and runs[-1][1] == is_g:
            runs[-1] = (runs[-1][0] + " " + s, is_g)
        else:
            runs.append((s, is_g))
    return runs


# --- низкоуровневые хелперы для OOXML -----------------------------------

def _set_cell_shading(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def _set_run_shading(run, hex_color: str):
    """Заливка текста внутри одного run произвольным hex-цветом — для
    точечной подсветки вопроса на грумминг прямо внутри ячейки таблицы,
    тем же светло-синим, что и у выносок-callout. python-docx не даёт
    готового метода для заливки на уровне run (только highlight_color
    с фиксированной палитрой), поэтому добавляем w:shd в rPr напрямую."""
    r_pr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    r_pr.append(shd)


def _set_cell_borders(cell, hex_color: str = BORDER_HEX, sz: str = "4", val: str = "single"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), hex_color)
        borders.append(el)
    tc_pr.append(borders)


def _set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tc_pr.append(mar)


def _no_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")
        borders.append(el)
    tbl_pr.append(borders)


def _shade_paragraph(paragraph, hex_color: str):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    p_pr.append(shd)


def _paragraph_left_border(paragraph, hex_color: str, sz: str = "18", space: str = "10"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), sz)
    left.set(qn("w:space"), space)
    left.set(qn("w:color"), hex_color)
    p_bdr.append(left)
    p_pr.append(p_bdr)


def _add_hyperlink(paragraph, url: str, text: str, color_hex: str = RULE_HEX, size_pt: float = 10.5):
    """Добавляет кликабельную гиперссылку внутрь paragraph (python-docx не
    даёт этого напрямую — нужен ручной OOXML w:hyperlink)."""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), color_hex)
    r_pr.append(color)

    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size_pt * 2)))
    r_pr.append(sz)

    r_fonts = OxmlElement("w:rFonts")
    r_fonts.set(qn("w:ascii"), BODY_FONT)
    r_fonts.set(qn("w:hAnsi"), BODY_FONT)
    r_pr.append(r_fonts)

    new_run.append(r_pr)
    t = OxmlElement("w:t")
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def _find_client_task_id(metadata: dict, field_name: str | None = None) -> str | None:
    """Ищет поле вида 'Идентификатор задачи клиента' в метаданных Трекера.
    Если field_name передан явно (из sections_example.json) — ищем точное
    совпадение (без учёта регистра), иначе — эвристика по словам
    'идентификатор' + 'клиент' в названии поля (у разных очередей Трекера
    кастомное поле может называться чуть по-разному)."""
    if not metadata:
        return None
    if field_name:
        for k, v in metadata.items():
            if k.strip().lower() == field_name.strip().lower() and v:
                return v
        return None
    for k, v in metadata.items():
        kl = k.lower()
        if "идентификатор" in kl and "клиент" in kl and v:
            return v
    return None


def _sanitize_filename_part(text: str, max_len: int = 80) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", text or "")
    text = re.sub(r"\s+", " ", text).strip(" .")
    if len(text) > max_len:
        text = text[:max_len].rstrip()
    return text or "без_названия"


def default_output_filename(data: dict, client_task_id_field: str | None = None) -> str:
    """<ID задачи клиента>_<Название задачи>.docx — если у задачи нет
    идентификатора в системе заказчика (поле пустое или это ссылка, а не
    короткий ID), используется ключ самой задачи в Yandex Tracker."""
    client_id = _find_client_task_id(data.get("metadata", {}), client_task_id_field)
    if client_id and not client_id.lower().startswith(("http://", "https://")):
        prefix = client_id
    else:
        prefix = data.get("key") or "task"
    title = data.get("title") or "Задача"
    return f"{_sanitize_filename_part(prefix)}_{_sanitize_filename_part(title)}.docx"


def _add_page_number_footer(section, label: str = ""):
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.text = ""
    if label:
        r = p.add_run(f"{label}  ·  ")
        r.font.size = Pt(10.5)
        r.font.color.rgb = MUTED
        r.font.name = BODY_FONT

    def field_run(field_code: str):
        run = p.add_run()
        run.font.size = Pt(10.5)
        run.font.color.rgb = MUTED
        run.font.name = BODY_FONT
        fld_begin = OxmlElement("w:fldChar")
        fld_begin.set(qn("w:fldCharType"), "begin")
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = field_code
        fld_end = OxmlElement("w:fldChar")
        fld_end.set(qn("w:fldCharType"), "end")
        run._r.append(fld_begin)
        run._r.append(instr)
        run._r.append(fld_end)

    field_run("PAGE")
    r = p.add_run(" / ")
    r.font.size = Pt(10.5)
    r.font.color.rgb = MUTED
    r.font.name = BODY_FONT
    field_run("NUMPAGES")


# --- выноски -------------------------------------------------------------

def _plain_runs(text: str) -> list[dict]:
    return [{"text": text, "bold": False, "italic": False}]


def _add_run_sequence(paragraph, runs_data: list[dict], font: str = BODY_FONT,
                       size: Pt | None = None, color: RGBColor | None = None):
    """Пишет в paragraph раны с сохранением жирного/курсива из исходной задачи."""
    for rd in runs_data:
        run = paragraph.add_run(rd["text"])
        run.font.name = font
        if size is not None:
            run.font.size = size
        if color is not None:
            run.font.color.rgb = color
        if rd.get("bold"):
            run.bold = True
        if rd.get("italic"):
            run.italic = True


def _add_callout(doc: Document, title: str, body_runs: list[dict],
                  bg_hex: str = CALLOUT_BG_HEX, border_hex: str = CALLOUT_BORDER_HEX,
                  title_color: RGBColor = None):
    p = doc.add_paragraph()
    label = title.strip().rstrip("!:").strip() or "Важно"
    r1 = p.add_run(f"{label}: ")
    r1.bold = True
    r1.font.name = BODY_FONT
    r1.font.size = Pt(11.5)
    r1.font.color.rgb = title_color or RGBColor(0x8A, 0x5A, 0x00)
    _add_run_sequence(p, body_runs, font=BODY_FONT, size=Pt(11.5))
    _shade_paragraph(p, bg_hex)
    _paragraph_left_border(p, border_hex)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.15)


# --- таблицы --------------------------------------------------------------

def _add_cell_line_runs(paragraph, line_runs: list[dict]):
    """line_runs: список {'text','bold','italic'} одной строки ячейки.
    Если строка содержит вопрос на грумминг — подсвечивает нужный фрагмент
    (та же логика, что и раньше), но тогда исходный bold/italic именно
    этой строки не сохраняется — редкое пересечение двух разных типов
    выделения, приоритет отдаётся видимости вопроса на грумминг. Если
    груминга нет — исходное форматирование (bold/italic) сохраняется как
    есть."""
    plain = "".join(r["text"] for r in line_runs)
    if GROOMING_TRIGGER_RE.search(plain):
        for chunk, is_grooming in _split_cell_grooming_runs(plain):
            run = paragraph.add_run(chunk)
            run.font.size = Pt(10.5)
            run.font.name = BODY_FONT
            if is_grooming:
                run.font.bold = True
                run.font.color.rgb = GROOMING_TEXT_COLOR
                _set_run_shading(run, GROOMING_BG_HEX)
        return
    for r in line_runs:
        run = paragraph.add_run(r["text"])
        run.font.size = Pt(10.5)
        run.font.name = BODY_FONT
        if r.get("bold"):
            run.bold = True
        if r.get("italic"):
            run.italic = True


def _cell_lines(cell_data) -> list[list[dict]]:
    """Нормализует значение ячейки к списку строк-ранов — cell_data обычно
    уже {'text','lines'} из tracker_parser, но на случай простой строки
    (например, если сборка вызвана напрямую с другими данными) — оборачивает."""
    if isinstance(cell_data, dict):
        lines = cell_data.get("lines")
        if lines is not None:
            return lines
        text = cell_data.get("text", "")
        return [[{"text": text, "bold": False, "italic": False}]] if text else []
    text = str(cell_data) if cell_data else ""
    return [[{"text": text, "bold": False, "italic": False}]] if text else []


def _cell_text(cell_data) -> str:
    if isinstance(cell_data, dict):
        return cell_data.get("text", "")
    return str(cell_data) if cell_data else ""


def _add_table_block(doc: Document, block: dict):
    header = block.get("header")
    rows = block.get("rows") or []
    n_cols = len(header) if header else (len(rows[0]) if rows else 0)
    if n_cols == 0:
        return

    n_rows = (1 if header else 0) + len(rows)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = True
    _no_table_borders(table)

    row_offset = 0
    if header:
        for j, cell_data in enumerate(header):
            cell = table.cell(0, j)
            cell.text = _cell_text(cell_data)
            run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else cell.paragraphs[0].add_run("")
            run.font.bold = True
            run.font.size = Pt(10.5)
            run.font.name = BODY_FONT
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _set_cell_shading(cell, TABLE_HEADER_HEX)
            _set_cell_margins(cell, top=80, bottom=80, left=140, right=140)
            _set_cell_borders(cell, TABLE_HEADER_HEX, "4")
        row_offset = 1

    for i, row in enumerate(rows):
        for j in range(n_cols):
            cell = table.cell(i + row_offset, j)
            cell_data = row[j] if j < len(row) else None
            lines = _cell_lines(cell_data)
            cell.text = ""
            if not lines:
                cell.paragraphs[0].paragraph_format.space_after = Pt(2)
            else:
                for li, line_runs in enumerate(lines):
                    p = cell.paragraphs[0] if li == 0 else cell.add_paragraph()
                    p.paragraph_format.space_after = Pt(2)
                    _add_cell_line_runs(p, line_runs)
            _set_cell_margins(cell, top=80, bottom=80, left=140, right=140)
            _set_cell_borders(cell, BORDER_HEX, "3")
            if (i % 2) == 1:
                _set_cell_shading(cell, ALT_ROW_HEX)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def is_business_case_heading(heading: str) -> bool:
    h = (heading or "").strip().lower()
    return any(alias in h for alias in BUSINESS_CASE_HEADINGS)


def has_business_case_section(sections: list[dict]) -> bool:
    return any(is_business_case_heading(s.get("heading", "")) for s in sections)


def has_grooming_mentions(sections: list[dict]) -> bool:
    """Ищет упоминания груминга где угодно — и в обычных блоках (после
    extract_grooming_blocks они уже выделены в отдельные параграфы), и
    внутри ячеек таблиц (там мы их не вычищаем — в структурированной
    таблице вырезать середину ячейки значит оставить дыру в данных,
    но для индикатора "есть упоминания" таблицы тоже должны учитываться)."""
    for s in sections:
        for b in s.get("content", []):
            btype = b.get("type")
            if btype in ("paragraph", "bullet", "number") and GROOMING_TRIGGER_RE.search(b.get("text", "")):
                return True
            if btype == "table":
                rows = ([b["header"]] if b.get("header") else []) + b.get("rows", [])
                for row in rows:
                    for cell in row:
                        if GROOMING_TRIGGER_RE.search(str(cell)):
                            return True
    return False


def _add_business_case_blocks(doc: Document, blocks: list[dict]):
    """Бизнес-кейс — самое важное для заказчика ("зачем это нужно бизнесу"),
    поэтому весь раздел визуально выделяется: каждый абзац/пункт списка
    получает мягкую заливку и акцентную полосу слева, как непрерывная
    выноска. Картинки (загруженные пользователем или плашка-заглушка)
    получают ту же рамку, чтобы не выпадать из общего блока — таблицы
    рендерятся как обычно (заливка целой таблицы визуально не нужна,
    у неё и так есть собственные границы и шапка)."""
    for block in blocks:
        btype = block.get("type")
        if btype == "table":
            _add_table_block(doc, block)
            for img in block.get("images", []):
                _add_image_block(doc, img, boxed=True)
            continue
        if btype == "image":
            _add_image_block(doc, block, boxed=True)
            continue

        if btype in ("bullet", "number"):
            style = "List Bullet" if btype == "bullet" else "List Number"
            try:
                p = doc.add_paragraph(style=style)
            except KeyError:
                p = doc.add_paragraph()
            _add_run_sequence(p, block.get("runs") or _plain_runs(block.get("text", "")), font=BODY_FONT, size=Pt(11.5))
            _shade_paragraph(p, BUSINESS_CASE_BG_HEX)
            p.paragraph_format.space_after = Pt(4)
            continue

        text = block.get("text", "").strip()
        if not text:
            continue
        p = doc.add_paragraph()
        _add_run_sequence(p, block.get("runs") or _plain_runs(text), font=BODY_FONT, size=Pt(11.5))
        _shade_paragraph(p, BUSINESS_CASE_BG_HEX)
        _paragraph_left_border(p, BUSINESS_CASE_BORDER_HEX, sz="18", space="10")
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Cm(0.15)


def _add_content_blocks(doc: Document, blocks: list[dict]):
    list_style_by_level = {
        0: ("List Bullet", "List Number"),
        1: ("List Bullet 2", "List Number 2"),
        2: ("List Bullet 3", "List Number 3"),
    }
    i = 0
    n = len(blocks)
    while i < n:
        block = blocks[i]
        btype = block.get("type")

        if btype == "table":
            _add_table_block(doc, block)
            for img in block.get("images", []):
                _add_image_block(doc, img)
            i += 1
            continue

        if btype == "image":
            _add_image_block(doc, block)
            i += 1
            continue

        if btype == "paragraph" and GROOMING_TRIGGER_RE.search(block.get("text", "")):
            _add_callout(
                doc, "Требуется грумминг", block.get("runs") or _plain_runs(block.get("text", "")),
                bg_hex=GROOMING_BG_HEX, border_hex=GROOMING_BORDER_HEX, title_color=GROOMING_TEXT_COLOR,
            )
            i += 1
            continue

        if btype == "paragraph" and CALLOUT_TRIGGER_RE.match(block.get("text", "")):
            title = block["text"]
            body_runs: list[dict] = []
            j = i + 1
            while (
                j < n
                and blocks[j].get("type") == "paragraph"
                and not CALLOUT_TRIGGER_RE.match(blocks[j].get("text", ""))
                and not GROOMING_TRIGGER_RE.search(blocks[j].get("text", ""))
            ):
                if body_runs:
                    body_runs.append({"text": " ", "bold": False, "italic": False})
                body_runs.extend(blocks[j].get("runs") or _plain_runs(blocks[j].get("text", "")))
                j += 1
            if body_runs:
                _add_callout(doc, title, body_runs)
                i = j
            else:
                i += 1
            continue

        if btype in ("bullet", "number"):
            level = min(block.get("level", 0), 2)
            bullet_style, number_style = list_style_by_level[level]
            style = number_style if btype == "number" else bullet_style
            try:
                p = doc.add_paragraph(style=style)
            except KeyError:
                p = doc.add_paragraph(style=list_style_by_level[0][0 if btype == "bullet" else 1])
                p.paragraph_format.left_indent = Cm(0.6 * (level + 1))
            _add_run_sequence(p, block.get("runs") or _plain_runs(block.get("text", "")),
                               font=BODY_FONT, size=Pt(11.5))
            p.paragraph_format.space_after = Pt(3)
        else:
            text = block.get("text", "").strip()
            if text:
                p = doc.add_paragraph()
                _add_run_sequence(p, block.get("runs") or _plain_runs(text), font=BODY_FONT, size=Pt(11.5))
                p.paragraph_format.space_after = Pt(8)
                p.paragraph_format.line_spacing = 1.2
        i += 1


# --- заголовки/шапка -------------------------------------------------------

def _add_title_band(doc: Document, key: str | None, title_text: str, logo_path: str | None,
                     company_name: str, client_task_id: str | None = None):
    section = doc.sections[0]
    content_width = section.page_width - section.left_margin - section.right_margin

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    table.columns[0].width = content_width
    _no_table_borders(table)
    cell = table.cell(0, 0)
    cell.width = content_width
    _set_cell_shading(cell, TITLE_BG_HEX)
    _set_cell_margins(cell, top=260, bottom=260, left=320, right=320)

    cell.paragraphs[0].text = ""
    first_par = cell.paragraphs[0]

    if logo_path or company_name:
        top_p = first_par
        if logo_path:
            run = top_p.add_run()
            run.add_picture(logo_path, height=Cm(1.0))
        if company_name:
            if logo_path:
                top_p.add_run("   ")
            r = top_p.add_run(company_name)
            r.font.size = Pt(10.5)
            r.font.name = BODY_FONT
            r.font.color.rgb = MUTED
        eyebrow_p = cell.add_paragraph()
    else:
        eyebrow_p = first_par

    if key:
        r = eyebrow_p.add_run(f"ЗАДАЧА {key}")
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.name = BODY_FONT
        r.font.color.rgb = ACCENT
        eyebrow_p.paragraph_format.space_after = Pt(2)

    if client_task_id:
        client_p = cell.add_paragraph()
        r = client_p.add_run("Задача в системе заказчика: ")
        r.font.size = Pt(10.5)
        r.font.name = BODY_FONT
        r.font.color.rgb = MUTED
        if client_task_id.startswith(("http://", "https://")):
            _add_hyperlink(client_p, client_task_id, client_task_id, color_hex=RULE_HEX, size_pt=10.5)
        else:
            rv = client_p.add_run(client_task_id)
            rv.font.size = Pt(10.5)
            rv.font.bold = True
            rv.font.name = BODY_FONT
            rv.font.color.rgb = ACCENT
        client_p.paragraph_format.space_after = Pt(2)

    title_p = cell.add_paragraph()
    r = title_p.add_run(title_text)
    r.font.size = Pt(24)
    r.font.bold = True
    r.font.name = HEADING_FONT
    r.font.color.rgb = ACCENT
    title_p.paragraph_format.space_after = Pt(0)
    title_p.paragraph_format.space_before = Pt(2)
    title_p.paragraph_format.line_spacing = 1.05

    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def _add_section_heading(doc: Document, text: str, color_hex: str = RULE_HEX, text_color: RGBColor = None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = HEADING_FONT
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = text_color or ACCENT
    _paragraph_left_border(p, color_hex, sz="20", space="10")
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(0.15)
    p.paragraph_format.keep_with_next = True
    return p


def _base_document_setup(doc: Document, footer_label: str = ""):
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(11.5)

    section = doc.sections[0]
    section.left_margin = Cm(2.3)
    section.right_margin = Cm(2.3)
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.different_first_page_header_footer = True

    _add_page_number_footer(section, label=footer_label)
    # первая страница — без колонтитула-лейбла (только номер), опрятнее
    first_footer_p = section.first_page_footer.paragraphs[0] if section.first_page_footer.paragraphs \
        else section.first_page_footer.add_paragraph()
    first_footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ============================================================================
# ПУБЛИЧНЫЕ ФУНКЦИИ
# ============================================================================

# Поля метаданных Трекера, которые по умолчанию показываем заказчику.
# По умолчанию — НИКАКИЕ: статус, даты создания/обновления, дедлайн,
# резолюция и подобные внутренние рабочие поля заказчику не нужны и не
# показываются. Если что-то из этого всё же нужно — передайте явный
# список через --sections-config (ключ metadata_fields), значения "Создано"
# /"Обновлено"/"Дедлайн" тоже поддерживаются.
DEFAULT_CLIENT_METADATA_FIELDS: list[str] = []

# Разделы, которые по умолчанию не показываем заказчику (внутренняя
# техническая информация), но всё ещё учитываем при проверке полноты задачи.
DEFAULT_CLIENT_EXCLUDED_SECTIONS = ["Техническая реализация"]


# Заголовки разделов, которые в клиентском документе показываются под другим,
# более уместным для заказчика названием — сопоставление ведётся по исходному
# заголовку из задачи (регистронезависимо), внутренняя логика (проверка
# обязательных разделов, исключения и т.п.) продолжает работать с исходным
# названием, меняется только то, что видит заказчик.
HEADING_DISPLAY_OVERRIDES = {
    "потребность клиента": "Потребность заказчика",
}


def _display_heading(heading: str) -> str:
    return HEADING_DISPLAY_OVERRIDES.get(heading.strip().lower(), heading)


def _add_image_block(doc: Document, block: dict, boxed: bool = False):
    """Сам пиксель картинки вшить нельзя (скриншоты в задаче — это ссылки на
    вложения Трекера, требующие авторизованной сессии), поэтому картинка
    становится компактной строкой-ссылкой на неё в Трекере.
    boxed=True — та же фиолетовая рамка+заливка, что у остального текста
    раздела "Бизнес-кейс" (чтобы картинка не выпадала из общего блока)."""
    def _box(p):
        if boxed:
            _shade_paragraph(p, BUSINESS_CASE_BG_HEX)
            _paragraph_left_border(p, BUSINESS_CASE_BORDER_HEX, sz="18", space="10")
            p.paragraph_format.left_indent = Cm(0.15)

    p = doc.add_paragraph()
    label = block.get("alt") or "Скриншот"
    run = p.add_run(f"🖼 {label}: ")
    run.italic = True
    run.font.name = BODY_FONT
    run.font.size = Pt(10.5)
    run.font.color.rgb = MUTED
    if block.get("url"):
        _add_hyperlink(p, block["url"], "открыть в Tracker", color_hex=RULE_HEX, size_pt=10.5)
    else:
        run2 = p.add_run("недоступен для предпросмотра — см. вложение в исходной задаче")
        run2.italic = True
        run2.font.name = BODY_FONT
        run2.font.size = Pt(10.5)
        run2.font.color.rgb = MUTED
    _box(p)
    if block.get("context"):
        p2 = doc.add_paragraph()
        run3 = p2.add_run(f"из строки таблицы: «{block['context']}»")
        run3.italic = True
        run3.font.name = BODY_FONT
        run3.font.size = Pt(9)
        run3.font.color.rgb = MUTED
        _box(p2)
        p2.paragraph_format.space_after = Pt(8)
    else:
        p.paragraph_format.space_after = Pt(8)


def _add_image_placeholder_block(doc: Document, block: dict):
    """Вариант A: надёжная плашка-слот вместо ссылки. Сам пиксель картинки
    по-прежнему не вшивается (нет доступа к сессии Tracker), но вместо
    незаметной строки-ссылки — чёткий пунктирный блок с контекстом, чтобы
    человек, открывший документ в Word, сразу видел: здесь нужно вставить
    скриншот вручную (Вставка → Рисунок), и знал, какой именно."""
    label = block.get("alt") or "Скриншот"
    section = doc.sections[0]
    content_width = section.page_width - section.left_margin - section.right_margin

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = content_width
    cell = table.cell(0, 0)
    cell.width = content_width
    _set_cell_shading(cell, "F7F9FA")
    _set_cell_borders(cell, "B9C2CB", sz="6", val="dashed")
    _set_cell_margins(cell, top=180, bottom=180, left=220, right=220)

    p0 = cell.paragraphs[0]
    r0 = p0.add_run(f"📷 Место для скриншота: {label}")
    r0.bold = True
    r0.font.name = BODY_FONT
    r0.font.size = Pt(10.5)
    r0.font.color.rgb = MUTED

    p1 = cell.add_paragraph()
    r1 = p1.add_run("Вставьте изображение сюда вручную: Вставка → Рисунок "
                     "(или перетащите файл прямо в это место)")
    r1.italic = True
    r1.font.name = BODY_FONT
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = MUTED

    if block.get("url"):
        p2 = cell.add_paragraph()
        rl = p2.add_run("Оригинал в задаче: ")
        rl.font.name = BODY_FONT
        rl.font.size = Pt(9.5)
        rl.font.color.rgb = MUTED
        _add_hyperlink(p2, block["url"], "открыть в Tracker", color_hex=RULE_HEX, size_pt=9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def build_docx_from_tracker(
    data: dict,
    output_path: str,
    missing_required: list[str] | None = None,
    logo_path: str | None = None,
    company_name: str = "",
    metadata_fields: list[str] | None = None,
    excluded_sections: list[str] | None = None,
    show_gaps: bool = True,
    client_task_id_field: str | None = None,
    show_screenshots: bool = True,
    screenshots_excluded_sections: list[str] | None = None,
    show_grooming: bool = True,
) -> str:
    """Собирает клиентский .docx из структуры, которую вернул tracker_parser.parse()
    (опционально пропущенной через rewriter.rewrite_sections).

    show_screenshots: показывать ли скриншоты вообще (по умолчанию да — они
    могут быть значимы для понимания задачи). Сама картинка не вшивается
    (нет доступа к сессии Трекера) — вставляется рабочая ссылка на неё.
    screenshots_excluded_sections: разделы, где скриншоты всё равно скрыть,
    даже при show_screenshots=True (например, слишком технический раздел).
    show_grooming: показывать ли выноски "Требуется грумминг" (по умолчанию
    да). При False вырезаются только блоки-параграфы/пункты списка — то,
    что extract_grooming_blocks уже успел вынести отдельным блоком при
    разборе; упоминания внутри ячеек таблиц не трогаются (см. docstring
    has_grooming_mentions)."""
    metadata_fields = metadata_fields if metadata_fields is not None else DEFAULT_CLIENT_METADATA_FIELDS
    excluded_sections = excluded_sections if excluded_sections is not None else DEFAULT_CLIENT_EXCLUDED_SECTIONS
    excluded_lower = {s.lower() for s in excluded_sections}
    screenshots_excluded_lower = {s.lower() for s in (screenshots_excluded_sections or [])}
    missing_required = missing_required or []

    doc = Document()
    key = data.get("key")
    title_text = data.get("title") or "Задача"
    client_task_id = _find_client_task_id(data.get("metadata", {}), client_task_id_field)
    _base_document_setup(doc, footer_label=f"{key}" if key else "")
    _add_title_band(doc, key, title_text, logo_path, company_name, client_task_id=client_task_id)

    # --- метаданные: только явный белый список полей, только то, что есть ---
    meta_rows = []
    date_fields = {"Создано": "created", "Обновлено": "updated", "Дедлайн": "deadline"}
    for field in metadata_fields:
        value = data.get(date_fields[field]) if field in date_fields else data.get("metadata", {}).get(field)
        if value:
            meta_rows.append((field, value))

    if meta_rows:
        table = doc.add_table(rows=len(meta_rows), cols=2)
        table.autofit = False
        table.columns[0].width = Cm(4)
        table.columns[1].width = Cm(11.5)
        _no_table_borders(table)
        for i, (k_, v_) in enumerate(meta_rows):
            cell_key = table.cell(i, 0)
            cell_key.text = k_
            cell_key.paragraphs[0].runs[0].font.bold = True
            cell_key.paragraphs[0].runs[0].font.size = Pt(10.5)
            cell_key.paragraphs[0].runs[0].font.name = BODY_FONT
            _set_cell_shading(cell_key, "F7F9FA")
            _set_cell_margins(cell_key, top=70, bottom=70, left=120, right=120)
            cell_val = table.cell(i, 1)
            cell_val.text = str(v_)
            cell_val.paragraphs[0].runs[0].font.size = Pt(10.5)
            cell_val.paragraphs[0].runs[0].font.name = BODY_FONT
            _set_cell_margins(cell_val, top=70, bottom=70, left=120, right=120)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # --- разделы задачи, в порядке документа ---
    for sec in data.get("sections", []):
        if sec["heading"].lower() in excluded_lower:
            continue

        is_business_case = is_business_case_heading(sec["heading"])
        heading_text = _display_heading(sec["heading"])
        if is_business_case:
            heading_text = "💼 " + heading_text
        _add_section_heading(
            doc, heading_text,
            color_hex=BUSINESS_CASE_BORDER_HEX if is_business_case else RULE_HEX,
            text_color=BUSINESS_CASE_TEXT_COLOR if is_business_case else ACCENT,
        )

        section_shows_screenshots = show_screenshots and sec["heading"].lower() not in screenshots_excluded_lower
        blocks = sec.get("content", [])
        if not section_shows_screenshots:
            blocks = [b for b in blocks if b.get("type") != "image"]
        if not show_grooming:
            blocks = [
                b for b in blocks
                if not (b.get("type") in ("paragraph", "bullet", "number")
                        and GROOMING_TRIGGER_RE.search(b.get("text", "")))
            ]
        if blocks:
            if is_business_case:
                _add_business_case_blocks(doc, blocks)
            else:
                _add_content_blocks(doc, blocks)
        else:
            p = doc.add_paragraph()
            run = p.add_run("Раздел присутствует в задаче, но не содержит текста.")
            run.italic = True
            run.font.name = BODY_FONT
            run.font.color.rgb = MUTED

    # --- рекомендация: в задаче нет раздела "Бизнес-кейс" ---
    if show_gaps and not has_business_case_section(data.get("sections", [])):
        _add_callout(
            doc, "❗ Рекомендация",
            _plain_runs(
                "В задаче нет раздела «Бизнес-кейс» — рекомендуем добавить его в Трекере, "
                "чтобы заказчику было понятно, какую функцию или ценность для бизнеса закрывает эта задача."
            ),
            bg_hex=BUSINESS_CASE_BG_HEX, border_hex=BUSINESS_CASE_BORDER_HEX, title_color=BUSINESS_CASE_TEXT_COLOR,
        )

    # --- отсутствующие обязательные разделы ---
    if show_gaps and missing_required:
        _add_section_heading(doc, "Требует уточнения", color_hex=WARN_BORDER_HEX, text_color=WARN)
        _add_callout(
            doc,
            "Обратите внимание",
            _plain_runs("В исходной задаче не найдены обязательные разделы: "
                        + ", ".join(_display_heading(h) for h in missing_required) + "."),
            bg_hex=WARN_BG_HEX, border_hex=WARN_BORDER_HEX, title_color=WARN,
        )

    doc.save(output_path)
    return output_path


def build_docx(
    data: dict,
    output_path: str,
    logo_path: str | None = None,
    company_name: str = "",
    show_gaps: bool = True,
) -> str:
    """Резервный путь для build_generic_pipeline (не-Tracker источники).
    data: JSON-словарь от analyzer.analyze() — content разделов там строка,
    не список блоков, поэтому рендерим его попроще (без настоящих таблиц)."""
    doc = Document()
    key = data.get("task_id")
    title_text = data.get("title") or "Задача"
    _base_document_setup(doc, footer_label=key or "")
    _add_title_band(doc, key, title_text, logo_path, company_name)

    meta_fields = [
        ("Исполнитель", data.get("assignee")),
        ("Статус", data.get("status")),
        ("Дата", data.get("date")),
    ]
    meta_fields = [(k, v) for k, v in meta_fields if v]
    if meta_fields:
        table = doc.add_table(rows=len(meta_fields), cols=2)
        table.autofit = False
        table.columns[0].width = Cm(4)
        table.columns[1].width = Cm(11.5)
        _no_table_borders(table)
        for i, (k_, v_) in enumerate(meta_fields):
            cell_key = table.cell(i, 0)
            cell_key.text = k_
            cell_key.paragraphs[0].runs[0].font.bold = True
            cell_key.paragraphs[0].runs[0].font.size = Pt(10.5)
            _set_cell_shading(cell_key, "F7F9FA")
            _set_cell_margins(cell_key, top=70, bottom=70, left=120, right=120)
            cell_val = table.cell(i, 1)
            cell_val.text = str(v_)
            cell_val.paragraphs[0].runs[0].font.size = Pt(10.5)
            _set_cell_margins(cell_val, top=70, bottom=70, left=120, right=120)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    for sec in data.get("sections", []):
        found = sec.get("found")
        content = sec.get("content")
        if not found and not show_gaps:
            continue

        _add_section_heading(doc, sec["name"])

        if found and content:
            for line in str(content).split("\n"):
                line = line.strip()
                if not line:
                    continue
                if line.startswith("- "):
                    p = doc.add_paragraph(line[2:], style="List Bullet")
                else:
                    p = doc.add_paragraph(line)
                for r in p.runs:
                    r.font.name = BODY_FONT
                    r.font.size = Pt(11.5)
        else:
            p = doc.add_paragraph()
            run = p.add_run("Не указано в исходной задаче")
            run.italic = True
            run.font.color.rgb = WARN

    extra = data.get("extra_section")
    if extra:
        _add_section_heading(doc, "Дополнительная информация")
        for line in str(extra).split("\n"):
            line = line.strip()
            if line:
                p = doc.add_paragraph(line)
                for r in p.runs:
                    r.font.name = BODY_FONT
                    r.font.size = Pt(11.5)

    doc.save(output_path)
    return output_path

"""
tracker_parser.py
Парсер .doc-экспорта из Yandex Tracker (реально это HTML с
Content-Type: application/vnd.ms-word).

Структура, которую генерирует Трекер, регулярна и одинакова для всех задач
одной очереди:
  1. h3 с [KEY] и заголовком задачи, рядом — "Создано:", "Обновлено:",
     иногда "Дедлайн:"
  2. одна-две таблицы служебных метаданных (Статус, Тип, Приоритет,
     Автор, Исполнитель, Проект, Оценка, ...) — набор колонок "Название: Значение"
  3. <td id="descriptionArea"> — само тело задачи, которое в этой команде
     всегда оформлено через h2-заголовки: "Потребность клиента",
     "Детали реализации", "Критерии приёмки" (Given/When/Then-таблица),
     "Техническая реализация" и т.п. Порядок и состав секций может
     отличаться от задачи к задаче.
  4. После — таблица "Комментарии" и тесткейсы — внутренняя информация,
     для заказчика не нужна.

Содержимое каждой секции возвращается не готовым текстом, а списком
СТРУКТУРНЫХ БЛОКОВ ({'type': 'paragraph'|'bullet'|'number', ...} или
{'type': 'table', 'header': [...], 'rows': [[...], ...]}), чтобы builder.py
мог отрисовать настоящую таблицу и настоящий список в .docx, а не свалить
всё в один абзац через "|".
"""
from __future__ import annotations

import re
from pathlib import Path

from bs4 import BeautifulSoup, NavigableString, Tag


def is_tracker_export(soup: BeautifulSoup) -> bool:
    """Быстрая проверка: похоже ли это на экспорт Yandex Tracker."""
    return soup.find(id="descriptionArea") is not None


def _clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _strip_leading_emoji(text: str) -> str:
    return re.sub(r"^[^\wа-яА-Я]+", "", text).strip().lower()


def _find_tracker_base_url(raw_html: str, key: str | None) -> str | None:
    """Ищет в самом экспортированном файле полный домен Трекера — Трекер
    сам вставляет пермалинк на текущую задачу вида 'https://DOMAIN/r/KEY'
    (например, в служебных ссылках), это единственный надёжный источник
    домена, раз сам .doc-файл его нигде явно не указывает."""
    if not key:
        return None
    m = re.search(r'https?://([^/"\s]+)/r/' + re.escape(key), raw_html)
    return f"https://{m.group(1)}" if m else None


# используется _render_blocks для превращения относительных src картинок
# в рабочие ссылки — выставляется на время одного parse() (см. ниже)
_CURRENT_BASE_URL: str | None = None
_IMAGE_ID_COUNTER: int = 0


# --------------------------------------------------- жирный/курсив (runs) --

def _extract_runs(node, bold: bool = False, italic: bool = False) -> list[dict]:
    """Рекурсивно собирает инлайн-содержимое узла в список сегментов —
    либо {'text','bold','italic'} (кусок текста), либо {'_image':True,
    'url','alt'} (картинка прямо посреди текста, например скриншот внутри
    абзаца юзерстори). Так жирный/курсив и картинки не теряются при сведении
    к обычному тексту, а порядок текст/картинка/текст сохраняется."""
    if isinstance(node, NavigableString):
        text = str(node)
        return [{"text": text, "bold": bold, "italic": italic}] if text else []
    if not isinstance(node, Tag):
        return []
    name = node.name
    if name in ("script", "style"):
        return []
    if name == "br":
        return [{"text": " ", "bold": bold, "italic": italic}]
    if name == "img":
        img_block = _image_block_from_tag(node)
        return [{"_image": True, **img_block}] if img_block else []
    new_bold = bold or name in ("strong", "b")
    new_italic = italic or name in ("em", "i")
    runs: list[dict] = []
    for child in node.children:
        runs.extend(_extract_runs(child, new_bold, new_italic))
    return runs


def _merge_runs(runs: list[dict]) -> list[dict]:
    """Схлопывает пробелы внутри каждого рана и склеивает соседние раны
    с одинаковым форматированием — без этого на стыках тегов появлялись бы
    лишние разрывы. Ведущий/конечный пробел всей последовательности обрезается,
    внутренние пробелы (между разно оформленными словами) сохраняются."""
    merged: list[dict] = []
    for r in runs:
        text = re.sub(r"[ \t\r\n]+", " ", r["text"])
        if not text:
            continue
        if merged and merged[-1]["bold"] == r["bold"] and merged[-1]["italic"] == r["italic"]:
            merged[-1]["text"] += text
        else:
            merged.append({"text": text, "bold": r["bold"], "italic": r["italic"]})
    if merged:
        merged[0]["text"] = merged[0]["text"].lstrip()
        merged[-1]["text"] = merged[-1]["text"].rstrip()
    return [r for r in merged if r["text"]]


def _image_block_from_tag(node) -> dict | None:
    global _IMAGE_ID_COUNTER
    src = (node.get("src") or "").strip()
    if not src:
        return None
    alt = _clean(node.get("alt") or "") or "Скриншот"
    if src.startswith(("http://", "https://")):
        abs_url = src
    elif _CURRENT_BASE_URL:
        abs_url = _CURRENT_BASE_URL.rstrip("/") + "/" + src.lstrip("/")
    else:
        abs_url = None
    _IMAGE_ID_COUNTER += 1
    return {"url": abs_url, "alt": alt, "id": f"img-{_IMAGE_ID_COUNTER}"}


def _paragraph_block(node) -> list[dict]:
    """Параграф/заголовок -> список блоков (обычно один, но если внутри
    есть картинка — текст до неё, сама картинка и текст после становятся
    отдельными последовательными блоками, чтобы картинка не потерялась)."""
    segments = _extract_runs(node)
    blocks: list[dict] = []
    current_runs: list[dict] = []

    def flush_text():
        merged = _merge_runs(current_runs)
        if merged:
            plain = "".join(r["text"] for r in merged)
            blocks.append({"type": "paragraph", "text": plain, "runs": merged})
        current_runs.clear()

    for seg in segments:
        if seg.get("_image"):
            flush_text()
            blocks.append({"type": "image", "url": seg.get("url"), "alt": seg.get("alt", "Скриншот"), "id": seg.get("id")})
        else:
            current_runs.append(seg)
    flush_text()
    return blocks


# ---------------------------------------------------------------- таблицы --

def _extract_cell_lines(td) -> list[list[dict]]:
    """Извлекает содержимое ячейки таблицы в виде списка строк (каждая
    строка — список ранов с сохранённым bold/italic). Раньше через
    get_text() всё схлопывалось в одну строку без форматирования — если
    в ячейке несколько <p> (частое явление в таблицах критериев приёмки,
    когда описание сценария состоит из нескольких абзацев), каждый теперь
    становится отдельной строкой при рендере, а не склеивается пробелом."""
    paragraphs = td.find_all("p", recursive=False)
    nodes = paragraphs if paragraphs else [td]
    lines: list[list[dict]] = []
    for node in nodes:
        raw_runs = [r for r in _extract_runs(node) if "text" in r]  # картинки собираются отдельно через _images_from_table
        runs = _merge_runs(raw_runs)
        if runs:
            lines.append(runs)
    return lines


def _cell_from_tag(td) -> dict:
    lines = _extract_cell_lines(td)
    plain = "\n".join("".join(r["text"] for r in line) for line in lines)
    return {"text": plain, "lines": lines}


def _plain_cell(text: str) -> dict:
    """Ячейка из простой строки (для служебных колонок вроде номера
    сценария, которых нет в исходном HTML) — без жирного/курсива."""
    text = text or ""
    lines = [[{"text": text, "bold": False, "italic": False}]] if text else []
    return {"text": text, "lines": lines}


def _images_from_table(table: Tag) -> list[dict]:
    """Картинки внутри ячеек таблицы (например, скриншот прямо в строке
    Given/When/Then) раньше молча терялись — get_text() их не видит.
    Собираем их отдельно, с контекстом (текст всей строки), чтобы потом
    показать после таблицы как отдельные слоты для вставки."""
    images: list[dict] = []
    for tr in table.find_all("tr"):
        cells_text = " | ".join(_clean(td.get_text(" ")) for td in tr.find_all(["td", "th"]))
        for img in tr.find_all("img"):
            block = _image_block_from_tag(img)
            if not block:
                continue
            context = cells_text
            if len(context) > 140:
                context = context[:140].rstrip() + "…"
            block["context"] = context
            images.append(block)
    return images


def _table_to_block(table: Tag) -> dict | None:
    rows = []
    for tr in table.find_all("tr"):
        cells = [_cell_from_tag(td) for td in tr.find_all(["td", "th"])]
        if any(c["text"] for c in cells):
            rows.append(cells)
    if not rows:
        return None
    header, body = rows[0], rows[1:]
    images = _images_from_table(table)
    if not body:
        # таблица из одной строки — трактуем как строку без шапки
        result = {"type": "table", "header": None, "rows": [header]}
    else:
        result = {"type": "table", "header": header, "rows": body}
    if images:
        result["images"] = images
    return result


# ------------------------------------------------------------------ списки --

def _list_to_blocks(list_tag: Tag, ordered: bool, level: int = 0) -> list[dict]:
    blocks: list[dict] = []
    for li in list_tag.find_all("li", recursive=False):
        own_blocks: list[dict] = []
        nested_lists: list[Tag] = []
        for child in li.children:
            if isinstance(child, Tag) and child.name in ("ul", "ol"):
                nested_lists.append(child)
                continue
            own_blocks.extend(_render_blocks(child))

        run_groups: list[dict] = []
        for b in own_blocks:
            if b.get("type") == "paragraph" and b.get("text"):
                if run_groups:
                    run_groups.append({"text": " ", "bold": False, "italic": False})
                run_groups.extend(b.get("runs") or [{"text": b["text"], "bold": False, "italic": False}])
        merged_runs = _merge_runs(run_groups)
        own_text = "".join(r["text"] for r in merged_runs)
        if own_text:
            blocks.append({
                "type": "number" if ordered else "bullet",
                "text": own_text,
                "level": level,
                "runs": merged_runs,
            })
        # если внутри <li> была таблица и т.п. — сохраняем как есть
        for b in own_blocks:
            if b.get("type") != "paragraph":
                blocks.append(b)

        for nested in nested_lists:
            blocks.extend(_list_to_blocks(nested, nested.name == "ol", level + 1))
    return blocks


# --------------------------------------------------------- общий рендерер --

def _render_blocks(node) -> list[dict]:
    """Рекурсивно превращает произвольный HTML-узел в список структурных
    блоков. Нужен из-за вложенных блоков Yandex Flavored Markdown —
    сворачиваемые '<details>/<summary>' и заметки '<div class="yfm-note">',
    где текст summary часто дублируется в 'yfm-note-title'."""
    if isinstance(node, NavigableString):
        t = _clean(str(node))
        return [{"type": "paragraph", "text": t}] if t else []
    if not isinstance(node, Tag):
        return []

    name = node.name
    if name in ("script", "style"):
        return []
    if name == "table":
        tb = _table_to_block(node)
        return [tb] if tb else []
    if name in ("ul", "ol"):
        return _list_to_blocks(node, name == "ol")
    if name == "br":
        return []
    if name == "img":
        img_block = _image_block_from_tag(node)
        return [{"type": "image", **img_block}] if img_block else []

    if name == "details":
        summary = node.find("summary", recursive=False)
        summary_text = _clean(summary.get_text(" ")) if summary else ""
        summary_norm = _strip_leading_emoji(summary_text)
        blocks = [{"type": "paragraph", "text": summary_text}] if summary_text else []
        for child in node.children:
            if child is summary:
                continue
            for b in _render_blocks(child):
                if (
                    b.get("type") == "paragraph"
                    and summary_norm
                    and _strip_leading_emoji(b.get("text", "")) == summary_norm
                ):
                    continue  # дублирующийся заголовок (yfm-note-title == summary)
                blocks.append(b)
        return blocks

    if name in ("div", "section", "blockquote", "span"):
        blocks = []
        for child in node.children:
            blocks.extend(_render_blocks(child))
        return blocks

    if name in ("p", "h1", "h2", "h3", "h4", "h5", "h6", "summary"):
        return _paragraph_block(node)

    return _paragraph_block(node)


# -------------------------------------------------- критерии приёмки (G/W/T) --

GROOMING_TRIGGER_RE = re.compile(r"грум{1,2}инг", re.IGNORECASE)
_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?])\s+|(?=\()")


def _split_sentences(text: str) -> list[str]:
    parts = _SENTENCE_SPLIT_RE.split(text)
    return [p for p in parts if p.strip()]


def _extract_grooming_from_text(text: str) -> tuple[str, list[str]]:
    """Разбивает текст на предложения (эвристика по точке/?/!), отделяет
    предложения с упоминанием груминга от остальных. Возвращает (текст без
    груминг-предложений, список вынесенных груминг-предложений).

    Вопрос на грумминг часто пишут в скобках одним заходом, но с несколькими
    предложениями внутри — только первое обычно содержит само слово
    "грумминг" ("(Вопрос на грумминг: ...? только если...?)"). Чтобы не
    резать такой вопрос пополам, отслеживаем баланс скобок: если
    груминг-предложение открывает скобку и не закрывает её сразу — все
    следующие предложения тоже считаются частью того же вопроса, пока
    скобка не закроется."""
    sentences = _split_sentences(text)
    if len(sentences) <= 1:
        return text, []  # нечего резать — весь текст либо груминг, либо нет
    kept: list[str] = []
    extracted: list[str] = []
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
        (extracted if is_g else kept).append(s.strip())
    return " ".join(kept).strip(), extracted


def extract_grooming_blocks(blocks: list[dict]) -> list[dict]:
    """Вопрос на груминг иногда пишут не отдельной строкой, а посреди
    длинного абзаца/пункта списка вперемешку с обычным текстом деталей
    реализации. Раньше такой абзац либо целиком превращался в выноску
    "Требуется грумминг" (теряя остальной текст из вида в выноске), либо
    (если груминг был лишь частью более длинного текста без разбиения на
    предложения) вообще не замечался. Здесь разбиваем такие блоки на
    предложения: груминг-предложения выносятся отдельным блоком сразу
    после исходного (дальше builder.py уже умеет превращать такой блок
    в цветную выноску по тому же ключевому слову), обычный текст
    остаётся на месте очищенным."""
    new_blocks: list[dict] = []
    for b in blocks:
        btype = b.get("type")
        text = b.get("text", "")
        if btype in ("paragraph", "bullet", "number") and text and GROOMING_TRIGGER_RE.search(text):
            cleaned, extracted = _extract_grooming_from_text(text)
            if extracted:
                if cleaned:
                    new_b = dict(b)
                    new_b["text"] = cleaned
                    new_b.pop("runs", None)  # раны считаны для исходного текста, после обрезки уже не точны
                    new_blocks.append(new_b)
                new_blocks.append({"type": "paragraph", "text": " ".join(extracted)})
                continue
        new_blocks.append(b)
    return new_blocks


def humanize_criteria_blocks(blocks: list[dict]) -> list[dict]:
    """Если в блоках есть Given/When/Then-таблица ('Дано'/'Когда'/'Тогда'),
    оставляет в ней только эти три колонки + номер сценария — остальное
    (например служебный ID тесткейса вроде 'ТК-37598-1') отбрасывается,
    это внутренняя ссылка, заказчику не нужна."""
    new_blocks = []
    for b in blocks:
        if b.get("type") == "table" and b.get("header"):
            header_norm = [c["text"].strip() for c in b["header"]]
            if {"Дано", "Когда", "Тогда"}.issubset(set(header_norm)):
                idx = {h: i for i, h in enumerate(header_norm)}
                new_rows = []
                for i, row in enumerate(b["rows"], 1):
                    def get(col):
                        j = idx.get(col)
                        return row[j] if j is not None and j < len(row) else _plain_cell("")
                    new_rows.append([_plain_cell(str(i)), get("Дано"), get("Когда"), get("Тогда")])
                new_table = {
                    "type": "table",
                    "header": [_plain_cell("#"), _plain_cell("Дано"), _plain_cell("Когда"), _plain_cell("Тогда")],
                    "rows": new_rows,
                }
                if b.get("images"):
                    new_table["images"] = b["images"]
                new_blocks.append(new_table)
                continue
        new_blocks.append(b)
    return new_blocks


# ------------------------------------------------------------------ заголовок --

def _parse_header(soup: BeautifulSoup) -> dict:
    h3 = soup.find("h3")
    out = {"key": None, "title": None, "created": None, "updated": None, "deadline": None}
    if not h3:
        return out

    lead_text = "".join(str(c) for c in h3.children if isinstance(c, NavigableString))
    m = re.search(r"\[([A-ZА-Я0-9]+-\d+)\]", lead_text)
    if m:
        out["key"] = m.group(1)

    link = h3.find("a")
    if link:
        out["title"] = _clean(link.get_text(" "))

    if not out["key"]:
        m = re.search(r"\[([A-ZА-Я0-9]+-\d+)\]", h3.get_text())
        if m:
            out["key"] = m.group(1)

    sub = h3.find(class_="subText")
    if sub:
        sub_text = _clean(sub.get_text(" "))
        labels = ["Создано", "Обновлено", "Дедлайн"]
        pattern = r"(" + "|".join(labels) + r"):\s*(.*?)(?=(?:" + "|".join(labels) + r"):|$)"
        for label_match, value_match in re.findall(pattern, sub_text):
            field = {"Создано": "created", "Обновлено": "updated", "Дедлайн": "deadline"}[label_match]
            out[field] = value_match.strip()

    return out


def _parse_metadata(soup: BeautifulSoup) -> dict:
    metadata: dict[str, str] = {}
    desc = soup.find(id="descriptionArea")
    for table in soup.find_all("table"):
        if desc and desc in table.parents:
            continue
        if table.find(id="descriptionArea"):
            continue
        for tr in table.find_all("tr"):
            cells = tr.find_all("td", recursive=False)
            i = 0
            while i < len(cells) - 1:
                label_b = cells[i].find("b")
                if label_b:
                    label = _clean(label_b.get_text(" ")).rstrip(":")
                    value = _clean(cells[i + 1].get_text(" "))
                    if label and value:
                        metadata[label] = value
                    i += 2
                else:
                    i += 1
    return metadata


_BUSINESS_CASE_LABEL_RE = re.compile(r"^(бизнес-кейс|бизнес кейс)", re.IGNORECASE)


def _is_business_case_label(text: str) -> bool:
    """Иногда «Бизнес-кейс» оформляют не заголовком, а сворачиваемым блоком
    (YFM <details><summary>Бизнес-кейс ...</summary>...</details>) — по виду
    как подпись/лейбл, а не заголовок раздела. Проверяем, что текст похож
    именно на такую подпись (короткий, начинается с «Бизнес-кейс»), а не
    просто где-то в длинном предложении упоминает эти слова."""
    t = (text or "").strip()
    return bool(_BUSINESS_CASE_LABEL_RE.match(t)) and len(t) <= 80


def _parse_description_sections(soup: BeautifulSoup) -> list[dict]:
    """descriptionArea -> список {'heading': str, 'content': [blocks]} в порядке документа."""
    desc = soup.find(id="descriptionArea")
    if not desc:
        return []

    sections: list[dict] = []
    current_heading = None
    current_blocks: list[dict] = []
    heading_already_emitted = False  # чтобы не плодить пустой дубль-хвост после вынесенного бизнес-кейса

    def flush():
        nonlocal heading_already_emitted
        if current_heading is not None:
            if heading_already_emitted and not current_blocks:
                return
            sections.append({"heading": current_heading, "content": current_blocks})
            heading_already_emitted = True

    for child in desc.children:
        if not isinstance(child, Tag):
            continue
        if child.name in ("h1", "h2", "h3", "h4"):
            flush()
            current_heading = _clean(child.get_text(" "))
            current_blocks = []
            heading_already_emitted = False
            continue

        if child.name == "details":
            summary = child.find("summary", recursive=False)
            summary_text = _clean(summary.get_text(" ")) if summary else ""
            if _is_business_case_label(summary_text):
                # автор оформил "Бизнес-кейс" сворачиваемым блоком, а не
                # заголовком — выносим его в отдельный раздел (тогда его
                # подхватит обычная детекция/подсветка бизнес-кейса), а
                # после закрытия details продолжаем текущий раздел как
                # ни в чём не бывало
                flush()
                bc_blocks: list[dict] = []
                for sub in child.children:
                    if sub is summary:
                        continue
                    bc_blocks.extend(_render_blocks(sub))
                sections.append({"heading": summary_text, "content": bc_blocks})
                current_blocks = []
                heading_already_emitted = True
                continue

        if current_heading is None:
            current_heading = "Описание"
            current_blocks = []
        current_blocks.extend(_render_blocks(child))
    flush()
    return sections


def parse(input_path: str) -> dict:
    global _CURRENT_BASE_URL, _IMAGE_ID_COUNTER
    path = Path(input_path)
    raw = path.read_text(encoding="utf-8", errors="ignore")
    soup = BeautifulSoup(raw, "html.parser")

    if not is_tracker_export(soup):
        raise ValueError("Файл не похож на экспорт Yandex Tracker (нет descriptionArea)")

    header = _parse_header(soup)
    _CURRENT_BASE_URL = _find_tracker_base_url(raw, header["key"])
    _IMAGE_ID_COUNTER = 0
    try:
        metadata = _parse_metadata(soup)
        sections = _parse_description_sections(soup)
        for sec in sections:
            if "риёмки" in sec["heading"].lower():
                sec["content"] = humanize_criteria_blocks(sec["content"])
            sec["content"] = extract_grooming_blocks(sec["content"])
    finally:
        _CURRENT_BASE_URL = None

    return {
        "key": header["key"],
        "title": header["title"],
        "created": header["created"],
        "updated": header["updated"],
        "deadline": header["deadline"],
        "metadata": metadata,
        "sections": sections,
    }

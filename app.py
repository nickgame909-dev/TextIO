"""
Учебный проект: форматирование Word (.docx) текстовыми запросами.

Схема работы:
    промт пользователя --> Gemini (google-genai) --> JSON-команды
    --> Python-функции на python-docx --> готовый .docx для скачивания
"""

import io
import json
import os
import re
import time
from collections import Counter

import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.parts.numbering import NumberingPart
from docx.shared import Cm, Pt, RGBColor
from google import genai
from google.genai import errors, types

# Google периодически снимает старые модели с доступа для новых пользователей.
# Если снова получите 404 NOT_FOUND — задайте другое имя в переменной окружения
# GEMINI_MODEL (актуальный список: https://ai.google.dev/gemini-api/docs/models).
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")

# Если основная модель перегружена (503) или недоступна (404), пробуем запасные по очереди.
MODEL_CANDIDATES = list(dict.fromkeys([MODEL_NAME, "gemini-3.7-flash", "gemini-3.5-flash-lite"]))
RETRY_CODES = {429, 500, 503, 504}  # временные ошибки: стоит повторить запрос
MAX_ATTEMPTS = 3  # попыток на каждую модель

# ---------------------------------------------------------------------------
# 1. СИСТЕМНЫЙ ПРОМТ: объясняем модели, в какой JSON переводить запрос
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """
You are a formatting assistant for Microsoft Word documents.
The user writes a request in natural language (any language, usually Russian).
Translate it into a JSON object with a list of commands. Output ONLY valid JSON,
no explanations, no markdown fences.

Format:
{"commands": [ {"action": "...", ...}, ... ]}

"target" object (used by set_alignment, set_font, set_paragraph_format, set_list, insert_page_break):
  {"type": "all" | "headings" | "body" | "paragraphs" | "contains" | "style",
   "indexes": [1, 2, -1],      // only for "paragraphs": 1-based, negative = from the end
   "text": "substring",        // only for "contains"
   "name": "Heading 1"}        // only for "style": exact paragraph style name
If the user does not specify which text, use {"type": "all"}.
"headings" = paragraphs with Heading/Title styles, "body" = all other non-empty paragraphs.

Available actions (include only the fields that need to change):

1. set_alignment: {"action":"set_alignment","target":{...},"align":"left|center|right|justify"}
2. set_font: {"action":"set_font","target":{...},"name":"Times New Roman","size_pt":14,
              "bold":true,"italic":false,"underline":false,"color":"FF0000"}
   color is a HEX string without '#': convert color names to HEX (red = FF0000, blue = 0000FF, black = 000000).
3. set_paragraph_format: {"action":"set_paragraph_format","target":{...},"line_spacing":1.5,
              "space_before_pt":0,"space_after_pt":6,"first_line_indent_cm":1.25}
4. set_columns: {"action":"set_columns","count":2,"spacing_cm":1.25,"line_between":false}
   (applies to the whole document; count=1 returns to one column)
5. set_header: {"action":"set_header","text":"Header text","align":"left|center|right|justify",
               "name":"Arial","size_pt":10,"bold":false,"italic":true,"color":"808080"}
6. set_footer: same fields as set_header. The text may contain the placeholder {PAGE}
   for the current page number and {NUMPAGES} for the total, e.g. "Page {PAGE} of {NUMPAGES}".
7. set_margins: {"action":"set_margins","top_cm":2,"bottom_cm":2,"left_cm":3,"right_cm":1.5}
8. set_orientation: {"action":"set_orientation","orientation":"portrait|landscape"}
9. set_page_size: {"action":"set_page_size","width_cm":21,"height_cm":29.7}  (A4 portrait; A4 landscape is 29.7 x 21)
10. set_list: {"action":"set_list","target":{...},"list_type":"bullet|numbered|none"}
11. set_table: {"action":"set_table","tables":"all","borders":true,"header_bold":true,
    "header_fill":"D9D9D9","header_repeat":true,"table_align":"left|center|right",
    "cell_align":"left|center|right|justify","name":"Arial","size_pt":11}
    ("tables": "all" or a list of 1-based table numbers)
12. insert_page_break: {"action":"insert_page_break","target":{...},"where":"before|after"}
13. insert_toc: {"action":"insert_toc","title":"Оглавление","levels":3,
    "position":"start|end","after_paragraph":2,"page_break_after":true}
14. replace_text: {"action":"replace_text","find":"old","replace":"new",
    "case_sensitive":false,"scope":"body|all"}  (scope "all" also searches headers/footers)
15. set_first_page: {"action":"set_first_page","different":true,"header_text":"","footer_text":""}
    Title page without a page number = set_footer with {PAGE} AND set_first_page with different=true.
16. set_style: {"action":"set_style","style":"Normal","name":"Times New Roman","size_pt":14,
    "bold":false,"color":"000000","align":"justify","line_spacing":1.5,"space_before_pt":0,
    "space_after_pt":0,"first_line_indent_cm":1.25,"left_indent_cm":0}
    Changes a paragraph style itself ("Normal" = body text, "Heading 1", "Title"...).
    Prefer it over set_font/set_paragraph_format when the user mentions styles or wants
    the whole document formatted consistently.

Rules:
- Use real numbers and booleans, not strings.
- One user request may need several commands; list them in order.
- If the request cannot be expressed with these actions, return {"commands": []}.
- I will give you an outline of the document (paragraph numbers, styles, first words).
  Use it to resolve references like "the third paragraph" or "the title".

Example request: "Сделай текст в две колонки, шрифт Arial 12, по ширине, внизу номер страницы по центру"
Example answer:
{"commands":[
 {"action":"set_columns","count":2,"spacing_cm":1.25,"line_between":false},
 {"action":"set_font","target":{"type":"all"},"name":"Arial","size_pt":12},
 {"action":"set_alignment","target":{"type":"all"},"align":"justify"},
 {"action":"set_footer","text":"{PAGE}","align":"center"}
]}
"""

ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

TWIPS_PER_CM = 567  # в XML Word расстояния между колонками хранятся в twips


# ---------------------------------------------------------------------------
# 2. РАБОТА С GEMINI
# ---------------------------------------------------------------------------
def get_api_key() -> str | None:
    """Ищем ключ: переменная окружения -> secrets.toml Streamlit."""
    key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not key:
        try:
            key = st.secrets.get("GEMINI_API_KEY")
        except Exception:
            key = None
    return key


def build_outline(doc: Document, limit: int = 40) -> str:
    """Краткий план документа, чтобы модель понимала, что такое «3-й абзац»."""
    lines = []
    for i, p in enumerate(doc.paragraphs, 1):
        if not p.text.strip():
            continue
        lines.append(f"{i}. [{p.style.name}] {p.text[:60]}")
        if len(lines) >= limit:
            lines.append("... (документ продолжается)")
            break
    header = f"Tables in document: {len(doc.tables)}\n"
    return header + ("\n".join(lines) or "(документ пустой)")


def parse_commands(raw: str) -> list[dict]:
    """Достаём список команд из ответа модели (с запасом на случай ```json)."""
    text = raw.strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if not match:
            raise ValueError("Модель вернула не JSON:\n" + raw)
        data = json.loads(match.group(0))
    if isinstance(data, dict):
        data = data.get("commands", [])
    if not isinstance(data, list):
        raise ValueError("В JSON нет списка commands")
    return [c for c in data if isinstance(c, dict)]


def ask_gemini(api_key: str, user_prompt: str, outline: str) -> tuple[list[dict], str]:
    """Возвращает (команды, имя модели, которая ответила)."""
    client = genai.Client(api_key=api_key)
    contents = f"Document outline:\n{outline}\n\nUser request:\n{user_prompt}"
    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        response_mime_type="application/json",  # просим вернуть чистый JSON
        temperature=0,  # меньше «творчества» — стабильнее команды
    )

    last_error = None
    for model in MODEL_CANDIDATES:
        for attempt in range(MAX_ATTEMPTS):
            try:
                response = client.models.generate_content(
                    model=model, contents=contents, config=config
                )
                return parse_commands(response.text or ""), model
            except errors.APIError as exc:
                last_error = exc
                if exc.code in RETRY_CODES:
                    time.sleep(2 ** attempt)  # пауза 1 с, 2 с, 4 с (экспоненциальная)
                    continue
                if exc.code == 404:
                    break  # модель недоступна — сразу берём следующую
                raise  # остальные ошибки (например, неверный ключ) повторять бессмысленно
    raise last_error


# ---------------------------------------------------------------------------
# 3. ПРИМЕНЕНИЕ КОМАНД К WORD (python-docx)
# ---------------------------------------------------------------------------
def all_paragraphs(doc: Document) -> list:
    """Абзацы основного текста + абзацы внутри таблиц (без дублей)."""
    result, seen = [], set()
    candidates = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                candidates.extend(cell.paragraphs)
    for p in candidates:
        if p._p not in seen:
            seen.add(p._p)
            result.append(p)
    return result


def is_heading(p) -> bool:
    name = (p.style.name or "").lower()
    return name.startswith(("heading", "title", "заголовок", "название"))


def select_paragraphs(doc: Document, target: dict | None) -> list:
    target = target or {"type": "all"}
    kind = target.get("type", "all")
    paragraphs = all_paragraphs(doc)

    if kind == "headings":
        return [p for p in paragraphs if is_heading(p)]
    if kind == "body":
        return [p for p in paragraphs if p.text.strip() and not is_heading(p)]
    if kind == "contains":
        needle = str(target.get("text", "")).lower()
        return [p for p in paragraphs if needle and needle in p.text.lower()]
    if kind == "style":
        name = str(target.get("name", "")).lower()
        return [p for p in paragraphs if p.style.name.lower() == name]
    if kind == "paragraphs":
        body = doc.paragraphs
        chosen = []
        for idx in target.get("indexes", []):
            idx = int(idx)
            pos = idx - 1 if idx > 0 else idx  # 1-based или с конца
            if -len(body) <= pos < len(body):
                chosen.append(body[pos])
        return chosen
    return paragraphs  # "all"


def set_run_font(run, spec: dict) -> None:
    """Применяет шрифт, размер, стиль и цвет к одному run (куску текста)."""
    name = spec.get("name") or spec.get("font_name")
    if name:
        run.font.name = name
        rfonts = run._element.rPr.rFonts
        rfonts.set(qn("w:eastAsia"), name)  # чтобы шрифт применился и к CJK/кириллице в Word
        rfonts.set(qn("w:cs"), name)
    if spec.get("size_pt"):
        run.font.size = Pt(float(spec["size_pt"]))
    for key in ("bold", "italic", "underline"):
        if spec.get(key) is not None:
            setattr(run.font, key, bool(spec[key]))
    color = str(spec.get("color", "")).lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", color):
        run.font.color.rgb = RGBColor.from_string(color.upper())


def cmd_set_alignment(doc, cmd):
    align = ALIGNMENTS.get(str(cmd.get("align", "")).lower())
    if align is None:
        raise ValueError(f"неизвестное выравнивание: {cmd.get('align')}")
    paragraphs = select_paragraphs(doc, cmd.get("target"))
    for p in paragraphs:
        p.alignment = align
    return f"выравнивание '{cmd['align']}' применено к {len(paragraphs)} абз."


def cmd_set_font(doc, cmd):
    paragraphs = select_paragraphs(doc, cmd.get("target"))
    for p in paragraphs:
        for run in p.runs:
            set_run_font(run, cmd)
    return f"шрифт изменён в {len(paragraphs)} абз."


def cmd_set_paragraph_format(doc, cmd):
    paragraphs = select_paragraphs(doc, cmd.get("target"))
    for p in paragraphs:
        fmt = p.paragraph_format
        if cmd.get("line_spacing") is not None:
            fmt.line_spacing = float(cmd["line_spacing"])
        if cmd.get("space_before_pt") is not None:
            fmt.space_before = Pt(float(cmd["space_before_pt"]))
        if cmd.get("space_after_pt") is not None:
            fmt.space_after = Pt(float(cmd["space_after_pt"]))
        if cmd.get("first_line_indent_cm") is not None:
            fmt.first_line_indent = Cm(float(cmd["first_line_indent_cm"]))
    return f"параметры абзаца изменены в {len(paragraphs)} абз."


def cmd_set_columns(doc, cmd):
    count = max(1, int(cmd.get("count", 1)))
    spacing = float(cmd.get("spacing_cm", 1.25))
    for section in doc.sections:
        sect_pr = section._sectPr
        cols = sect_pr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            # w:cols должен стоять в sectPr до перечисленных ниже элементов
            after = ["w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
                     "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid"]
            for tag in after:
                el = sect_pr.find(qn(tag))
                if el is not None:
                    el.addprevious(cols)
                    break
            else:
                sect_pr.append(cols)
        for child in list(cols):  # убираем старые настройки отдельных колонок
            cols.remove(child)
        cols.set(qn("w:num"), str(count))
        cols.set(qn("w:space"), str(int(spacing * TWIPS_PER_CM)))
        cols.set(qn("w:equalWidth"), "1")
        cols.set(qn("w:sep"), "1" if cmd.get("line_between") else "0")
    return f"колонок: {count}"


def add_field(paragraph, instruction: str, placeholder_text: str = "1"):
    """Вставляет поле Word (PAGE, NUMPAGES, TOC) — номер страницы обновляется сам."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = placeholder_text
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, placeholder, end])
    return run


def fill_header_footer(container, cmd: dict) -> None:
    """container — это section.header или section.footer."""
    container.is_linked_to_previous = False  # отвязываем от предыдущего раздела
    paragraphs = container.paragraphs
    paragraph = paragraphs[0]
    for p in paragraphs:
        p.clear()
    text = str(cmd.get("text", ""))
    for part in re.split(r"(\{PAGE\}|\{NUMPAGES\})", text):
        if part == "{PAGE}":
            run = add_field(paragraph, "PAGE")
        elif part == "{NUMPAGES}":
            run = add_field(paragraph, "NUMPAGES")
        elif part:
            run = paragraph.add_run(part)
        else:
            continue
        set_run_font(run, cmd)
    align = ALIGNMENTS.get(str(cmd.get("align", "center")).lower())
    if align is not None:
        paragraph.alignment = align


def cmd_set_header(doc, cmd):
    for section in doc.sections:
        fill_header_footer(section.header, cmd)
    return f"верхний колонтитул: «{cmd.get('text', '')}»"


def cmd_set_footer(doc, cmd):
    for section in doc.sections:
        fill_header_footer(section.footer, cmd)
    return f"нижний колонтитул: «{cmd.get('text', '')}»"


def cmd_set_margins(doc, cmd):
    for section in doc.sections:
        if cmd.get("top_cm") is not None:
            section.top_margin = Cm(float(cmd["top_cm"]))
        if cmd.get("bottom_cm") is not None:
            section.bottom_margin = Cm(float(cmd["bottom_cm"]))
        if cmd.get("left_cm") is not None:
            section.left_margin = Cm(float(cmd["left_cm"]))
        if cmd.get("right_cm") is not None:
            section.right_margin = Cm(float(cmd["right_cm"]))
    return "поля страницы изменены"


def cmd_set_orientation(doc, cmd):
    landscape = str(cmd.get("orientation", "")).lower() == "landscape"
    for section in doc.sections:
        width, height = section.page_width, section.page_height
        if landscape and width < height or not landscape and width > height:
            section.page_width, section.page_height = height, width
        section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    return "ориентация: " + ("альбомная" if landscape else "книжная")


# ---------------------------------------------------------------------------
# 3.1 ФОРМАТИРОВАНИЕ ПО ОБРАЗЦУ
# Идея: «читаем» оформление документа-образца и превращаем его в ТЕ ЖЕ JSON-команды,
# которые раньше давал Gemini. Дальше работает уже знакомый apply_commands().
# ---------------------------------------------------------------------------
ALIGN_NAMES = {value: name for name, value in ALIGNMENTS.items()}


def docdefault(doc: Document, xpath: str):
    """Значение из «настроек по умолчанию» документа (w:docDefaults в styles.xml)."""
    values = doc.styles.element.xpath(xpath)
    return str(values[0]) if values else None


def _font_value(font, attr: str):
    return font.color.rgb if attr == "color" else getattr(font, attr)


def resolve_font_attr(doc: Document, run, paragraph, attr: str):
    """
    Настоящее значение свойства шрифта. В Word оно часто не записано в самом тексте,
    а наследуется: run -> стиль абзаца -> родительские стили -> настройки по умолчанию.
    """
    value = _font_value(run.font, attr)
    if value is not None:
        return value
    style = paragraph.style
    while style is not None:
        value = _font_value(style.font, attr)
        if value is not None:
            return value
        style = style.base_style
    if attr == "name":
        return docdefault(doc, "w:docDefaults/w:rPrDefault/w:rPr/w:rFonts/@w:ascii")
    if attr == "size":
        half_points = docdefault(doc, "w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
        return Pt(int(half_points) / 2) if half_points else None
    return None


def resolve_pf(paragraph, attr: str):
    """То же наследование, но для свойств абзаца (выравнивание, интервалы, отступы)."""
    value = getattr(paragraph.paragraph_format, attr)
    if value is not None:
        return value
    style = paragraph.style
    while style is not None:
        value = getattr(style.paragraph_format, attr)
        if value is not None:
            return value
        style = style.base_style
    return None


def dominant_format(doc: Document, paragraphs: list) -> dict:
    """
    Находит «самое типичное» оформление группы абзацев: у каждого свойства выбирается
    значение, которое покрывает больше всего символов текста.
    """
    font_counts = {key: Counter() for key in ("name", "size_pt", "bold", "italic", "color")}
    para_keys = ("align", "line_spacing", "space_before_pt", "space_after_pt", "first_line_indent_cm")
    para_counts = {key: Counter() for key in para_keys}

    for p in paragraphs:
        weight = len(p.text.strip())
        if not weight:
            continue
        for run in p.runs:
            w = len(run.text.strip())
            if not w:
                continue
            name = resolve_font_attr(doc, run, p, "name")
            if name:
                font_counts["name"][str(name)] += w
            size = resolve_font_attr(doc, run, p, "size")
            if size is not None:
                font_counts["size_pt"][round(size.pt * 2) / 2] += w
            font_counts["bold"][bool(resolve_font_attr(doc, run, p, "bold"))] += w
            font_counts["italic"][bool(resolve_font_attr(doc, run, p, "italic"))] += w
            color = resolve_font_attr(doc, run, p, "color")
            if color is not None:
                font_counts["color"][str(color)] += w

        para_counts["align"][ALIGN_NAMES.get(resolve_pf(p, "alignment"), "left")] += weight
        spacing = resolve_pf(p, "line_spacing")
        if isinstance(spacing, float):  # float = множитель (1.5), иначе точное значение в пунктах
            para_counts["line_spacing"][round(spacing, 2)] += weight
        for key, attr, unit in (("space_before_pt", "space_before", "pt"),
                                ("space_after_pt", "space_after", "pt"),
                                ("first_line_indent_cm", "first_line_indent", "cm")):
            value = resolve_pf(p, attr)
            if value is not None:
                para_counts[key][round(getattr(value, unit), 2)] += weight

    def top(counter):
        return counter.most_common(1)[0][0] if counter else None

    result = {key: top(counter) for key, counter in font_counts.items()}
    result.update({key: top(counter) for key, counter in para_counts.items()})
    return result


def commands_from_format(fmt: dict, target: dict) -> list[dict]:
    """Превращает найденное оформление в JSON-команды для заданной группы абзацев."""
    commands = []
    font = {"action": "set_font", "target": target}
    for key in ("name", "size_pt", "bold", "italic", "color"):
        if fmt[key] is not None:
            font[key] = fmt[key]
    if len(font) > 2:
        commands.append(font)
    if fmt["align"]:
        commands.append({"action": "set_alignment", "target": target, "align": fmt["align"]})
    para = {"action": "set_paragraph_format", "target": target}
    for key in ("line_spacing", "space_before_pt", "space_after_pt", "first_line_indent_cm"):
        if fmt[key] is not None:
            para[key] = fmt[key]
    if len(para) > 2:
        commands.append(para)
    return commands


def field_placeholder(instruction: str) -> str:
    instruction = instruction.upper()
    if "NUMPAGES" in instruction:  # проверяем первым: слово NUMPAGES содержит PAGE
        return "{NUMPAGES}"
    if "PAGE" in instruction:
        return "{PAGE}"
    return ""


def paragraph_text_with_fields(paragraph) -> str:
    """Текст абзаца, где поля номера страницы заменены на {PAGE} / {NUMPAGES}."""
    parts, skip, in_result = [], set(), False
    for el in paragraph._p.iter():
        if el in skip:
            continue
        if el.tag == qn("w:fldSimple"):  # простое поле одним тегом
            parts.append(field_placeholder(el.get(qn("w:instr")) or ""))
            skip.update(el.iter(qn("w:t")))
        elif el.tag == qn("w:fldChar"):  # сложное поле: begin ... separate ... end
            kind = el.get(qn("w:fldCharType"))
            if kind == "separate":
                in_result = True  # дальше идёт сохранённый результат (например «1») — пропускаем
            elif kind == "end":
                in_result = False
        elif el.tag == qn("w:instrText"):
            parts.append(field_placeholder(el.text or ""))
        elif el.tag == qn("w:t") and not in_result:
            parts.append(el.text or "")
    return "".join(parts)


def header_footer_command(example: Document, action: str, container, copy_text: bool):
    if container.is_linked_to_previous:  # в образце такого колонтитула нет
        return None
    for p in container.paragraphs:
        text = paragraph_text_with_fields(p).strip()
        if not text:
            continue
        has_page_number = "{PAGE}" in text or "{NUMPAGES}" in text
        if not has_page_number and not copy_text:
            return None  # обычный текст колонтитула не копируем, если пользователь не просил
        fmt = dominant_format(example, [p])
        cmd = {"action": action, "text": text, "align": fmt["align"] or "left"}
        for key in ("name", "size_pt", "bold", "italic", "color"):
            if fmt[key] is not None:
                cmd[key] = fmt[key]
        return cmd
    return None


def extract_format_commands(example: Document, target: Document, copy_hf_text: bool = True) -> list[dict]:
    """Главная функция: документ-образец -> список JSON-команд."""
    commands = []
    section = example.sections[0]

    # Страница: размер и поля
    if section.page_width and section.page_height:
        commands.append({"action": "set_page_size",
                         "width_cm": round(section.page_width.cm, 2),
                         "height_cm": round(section.page_height.cm, 2)})
    margins = {}
    for key, attr in (("top_cm", "top_margin"), ("bottom_cm", "bottom_margin"),
                      ("left_cm", "left_margin"), ("right_cm", "right_margin")):
        value = getattr(section, attr)
        if value is not None:
            margins[key] = round(value.cm, 2)
    if margins:
        commands.append({"action": "set_margins", **margins})

    # Колонки читаем прямо из XML (тег w:cols)
    count, spacing, line = 1, 1.25, False
    cols = section._sectPr.find(qn("w:cols"))
    if cols is not None:
        count = int(cols.get(qn("w:num"), 1))
        spacing = round(int(cols.get(qn("w:space"), 708)) / TWIPS_PER_CM, 2)
        line = cols.get(qn("w:sep")) in ("1", "true", "on")
    commands.append({"action": "set_columns", "count": count, "spacing_cm": spacing, "line_between": line})

    # Колонтитулы
    for action, container in (("set_header", section.header), ("set_footer", section.footer)):
        cmd = header_footer_command(example, action, container, copy_hf_text)
        if cmd:
            commands.append(cmd)

    # Основной текст
    body = [p for p in example.paragraphs if p.text.strip() and not is_heading(p)]
    commands += commands_from_format(dominant_format(example, body), {"type": "body"})

    # Заголовки: отдельно для каждого стиля (Heading 1, Heading 2...), если он есть в целевом файле
    target_styles = {p.style.name for p in all_paragraphs(target)}
    groups: dict[str, list] = {}
    for p in example.paragraphs:
        if p.text.strip() and is_heading(p) and p.style.name in target_styles:
            groups.setdefault(p.style.name, []).append(p)
    for style_name, paragraphs in groups.items():
        commands += commands_from_format(
            dominant_format(example, paragraphs), {"type": "style", "name": style_name}
        )
    return commands


def cmd_set_page_size(doc, cmd):
    width, height = Cm(float(cmd["width_cm"])), Cm(float(cmd["height_cm"]))
    for section in doc.sections:
        section.page_width, section.page_height = width, height
        section.orientation = WD_ORIENT.LANDSCAPE if width > height else WD_ORIENT.PORTRAIT
    return f"размер страницы {cmd['width_cm']}×{cmd['height_cm']} см"


# ---------------------------------------------------------------------------
# 3.3 НОВЫЕ КОМАНДЫ: списки, таблицы, разрывы страниц, оглавление,
#     замена текста, колонтитулы первой страницы, стили
# ---------------------------------------------------------------------------
def insert_before_any(parent, new_el, successor_tags: list[str]) -> None:
    """
    Вставляет new_el перед первым найденным элементом из successor_tags.
    Word строго следит за порядком тегов внутри XML, поэтому просто append() часто ломает файл.
    """
    for tag in successor_tags:
        el = parent.find(qn(tag))
        if el is not None:
            el.addprevious(new_el)
            return
    parent.append(new_el)


# ----- 1. Списки (нумерованные и маркированные) -----------------------------
def numbering_element(doc: Document):
    """XML-корень numbering.xml. Если в документе такой части нет, создаём пустую."""
    try:
        return doc.part.numbering_part.element
    except (KeyError, NotImplementedError):
        pass
    part = NumberingPart.load(
        partname=PackURI("/word/numbering.xml"),
        content_type=CT.WML_NUMBERING,
        package=doc.part.package,
        blob=f"<w:numbering {nsdecls('w')}/>".encode(),
    )
    doc.part.relate_to(part, RT.NUMBERING)
    return part.element


def new_list_definition(doc: Document, kind: str) -> int:
    """Создаёт определение списка и возвращает numId (у каждой команды свой: нумерация с единицы)."""
    numbering = numbering_element(doc)
    abstract_ids = [int(e.get(qn("w:abstractNumId"))) for e in numbering.findall(qn("w:abstractNum"))]
    num_ids = [int(e.get(qn("w:numId"))) for e in numbering.findall(qn("w:num"))]
    abstract_id = max(abstract_ids, default=0) + 1
    num_id = max(num_ids, default=0) + 1
    fmt, text = ("bullet", "•") if kind == "bullet" else ("decimal", "%1.")
    abstract = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>'
        f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
        f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        f'</w:abstractNum>'
    )
    # по схеме все abstractNum идут раньше всех num
    insert_before_any(numbering, abstract, ["w:num"])
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'
    ))
    return num_id


def cmd_set_list(doc, cmd):
    kind = str(cmd.get("list_type", "bullet")).lower()
    if kind not in ("bullet", "numbered", "none"):
        raise ValueError(f"неизвестный тип списка: {kind}")
    paragraphs = [p for p in select_paragraphs(doc, cmd.get("target")) if p.text.strip()]
    if kind == "none":
        for p in paragraphs:
            pPr = p._p.pPr
            if pPr is not None and pPr.numPr is not None:
                pPr.remove(pPr.numPr)
        return f"список убран у {len(paragraphs)} абз."
    num_id = new_list_definition(doc, kind)
    for p in paragraphs:
        numPr = p._p.get_or_add_pPr().get_or_add_numPr()
        numPr.get_or_add_ilvl().val = 0
        numPr.get_or_add_numId().val = num_id
        p.paragraph_format.first_line_indent = None  # красная строка мешала бы отступу списка
        p.paragraph_format.left_indent = None
    return f"{'нумерованный' if kind == 'numbered' else 'маркированный'} список: {len(paragraphs)} пункт(ов)"


# ----- 2. Таблицы -----------------------------------------------------------
def unique_cells(cells) -> list:
    """Объединённые ячейки python-docx возвращает несколько раз; оставляем по одной."""
    seen, result = set(), []
    for cell in cells:
        if cell._tc not in seen:
            seen.add(cell._tc)
            result.append(cell)
    return result


def set_table_borders(table, on: bool) -> None:
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single" if on else "nil")
        if on:
            el.set(qn("w:sz"), "4")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "auto")
        borders.append(el)
    insert_before_any(tblPr, borders, ["w:shd", "w:tblLayout", "w:tblCellMar",
                                       "w:tblLook", "w:tblCaption", "w:tblDescription"])


def shade_cell(cell, fill_hex: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex.upper())
    insert_before_any(tcPr, shd, ["w:noWrap", "w:tcMar", "w:textDirection",
                                  "w:tcFitText", "w:vAlign", "w:hideMark"])


def cmd_set_table(doc, cmd):
    tables = doc.tables
    if not tables:
        return "в документе нет таблиц"
    selection = cmd.get("tables", "all")
    if selection in (None, "all"):
        chosen = list(tables)
    else:
        indexes = selection if isinstance(selection, list) else [selection]
        chosen = [tables[int(i) - 1] for i in indexes if 0 < int(i) <= len(tables)]

    has_font = any(cmd.get(k) is not None for k in ("name", "size_pt", "bold", "italic", "underline", "color"))
    cell_align = ALIGNMENTS.get(str(cmd.get("cell_align", "")).lower())
    table_align = {"left": WD_TABLE_ALIGNMENT.LEFT, "center": WD_TABLE_ALIGNMENT.CENTER,
                   "right": WD_TABLE_ALIGNMENT.RIGHT}.get(str(cmd.get("table_align", "")).lower())
    fill = str(cmd.get("header_fill", "")).lstrip("#")

    for table in chosen:
        if cmd.get("borders") is not None:
            set_table_borders(table, bool(cmd["borders"]))
        if table_align is not None:
            table.alignment = table_align
        cells = unique_cells(c for row in table.rows for c in row.cells)
        for cell in cells:
            for p in cell.paragraphs:
                if cell_align is not None:
                    p.alignment = cell_align
                if cmd.get("reset_indent", True):  # красная строка в ячейках почти всегда лишняя
                    p.paragraph_format.first_line_indent = Cm(0)
                if has_font:
                    for run in p.runs:
                        set_run_font(run, cmd)
        if table.rows:
            header_cells = unique_cells(table.rows[0].cells)
            if cmd.get("header_bold") is not None:
                for cell in header_cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.bold = bool(cmd["header_bold"])
            if re.fullmatch(r"[0-9a-fA-F]{6}", fill):
                for cell in header_cells:
                    shade_cell(cell, fill)
            if cmd.get("header_repeat"):  # шапка повторяется на каждой странице
                trPr = table.rows[0]._tr.get_or_add_trPr()
                if trPr.find(qn("w:tblHeader")) is None:
                    trPr.append(OxmlElement("w:tblHeader"))
    return f"оформлено таблиц: {len(chosen)}"


# ----- 3. Разрывы страниц ---------------------------------------------------
def cmd_insert_page_break(doc, cmd):
    where = str(cmd.get("where", "before")).lower()
    body = list(doc.paragraphs)
    body_ids = {p._p for p in body}
    # разрыв имеет смысл только для абзацев основного текста (не внутри таблиц)
    targets = [p for p in select_paragraphs(doc, cmd.get("target")) if p._p in body_ids]
    count = 0
    for p in targets:
        if where == "after":
            i = next(k for k, q in enumerate(body) if q._p is p._p)
            if i + 1 < len(body):
                body[i + 1].paragraph_format.page_break_before = True
            else:
                p.add_run().add_break(WD_BREAK.PAGE)
        else:
            p.paragraph_format.page_break_before = True
        count += 1
    return f"разрыв страницы {'после' if where == 'after' else 'перед'}: {count} абз."


# ----- 4. Оглавление --------------------------------------------------------
SETTINGS_AFTER_UPDATE_FIELDS = [
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars", "w:rsids",
    "w:themeFontLang", "w:clrSchemeMapping", "w:shapeDefaults", "w:decimalSymbol", "w:listSeparator",
]


def request_field_update(doc: Document) -> None:
    """Просит Word обновить поля (оглавление, номера страниц) при открытии файла."""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        insert_before_any(settings, el, SETTINGS_AFTER_UPDATE_FIELDS)


def cmd_insert_toc(doc, cmd):
    levels = min(9, max(1, int(cmd.get("levels", 3))))
    title = str(cmd.get("title", "Оглавление"))
    body = list(doc.paragraphs)

    title_p = None
    if title:
        title_p = doc.add_paragraph()
        run = title_p.add_run(title)
        run.bold = True
        run.font.size = Pt(16)
        title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_p = doc.add_paragraph()
    add_field(toc_p, f'TOC \\o "1-{levels}" \\h \\z \\u',
              placeholder_text="Оглавление появится после обновления полей (Ctrl+A, затем F9).")
    new_paragraphs = [p for p in (title_p, toc_p) if p is not None]

    position = str(cmd.get("position", "start")).lower()
    after = cmd.get("after_paragraph")
    if body and after is not None and 1 <= int(after) <= len(body):
        anchor, before = body[int(after) - 1]._p, False
    elif body and position != "end":
        anchor, before = body[0]._p, True
    else:
        anchor = None  # «в конец»: add_paragraph уже поставил абзацы куда нужно
    if anchor is not None:
        for p in (new_paragraphs if before else reversed(new_paragraphs)):
            anchor.addprevious(p._p) if before else anchor.addnext(p._p)

    if cmd.get("page_break_after"):
        toc_p.add_run().add_break(WD_BREAK.PAGE)
    request_field_update(doc)
    return f"оглавление до уровня {levels} (в Word нажмите «Обновить поле» или F9)"


# ----- 5. Замена текста -----------------------------------------------------
def replace_in_paragraph(paragraph, find: str, repl: str, case_sensitive: bool) -> int:
    """
    Замена, которая переживает разбиение текста на несколько run'ов
    (в Word одно слово часто лежит в 2-3 кусках из-за проверки орфографии и правок).
    Замена записывается в первый затронутый run, остальные только укорачиваются.
    """
    runs = paragraph.runs
    full = "".join(r.text for r in runs)
    pattern = re.compile(re.escape(find), 0 if case_sensitive else re.IGNORECASE)
    matches = list(pattern.finditer(full))
    for m in reversed(matches):  # с конца, чтобы не сбивались позиции предыдущих совпадений
        start, end = m.span()
        pos, first = 0, True
        for r in runs:
            text = r.text
            r_start, r_end = pos, pos + len(text)
            pos = r_end
            if not text or r_end <= start or r_start >= end:
                continue  # пустые run'ы (картинки, поля) не трогаем
            a, b = max(start, r_start) - r_start, min(end, r_end) - r_start
            r.text = text[:a] + (repl if first else "") + text[b:]
            first = False
    return len(matches)


def header_footer_paragraphs(doc: Document) -> list:
    result = []
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                     s.even_page_header, s.even_page_footer):
            if not part.is_linked_to_previous:
                result.extend(part.paragraphs)
    return result


def cmd_replace_text(doc, cmd):
    find = str(cmd.get("find", ""))
    if not find:
        raise ValueError("не указан текст для поиска (find)")
    repl = str(cmd.get("replace", ""))
    case = bool(cmd.get("case_sensitive", False))
    paragraphs = all_paragraphs(doc)
    if str(cmd.get("scope", "body")).lower() == "all":
        paragraphs += header_footer_paragraphs(doc)
    total = sum(replace_in_paragraph(p, find, repl, case) for p in paragraphs)
    return f"«{find}» → «{repl}»: {total} замен(ы)"


# ----- 6. Колонтитулы первой страницы ---------------------------------------
def cmd_set_first_page(doc, cmd):
    different = bool(cmd.get("different", True))
    for section in doc.sections:
        section.different_first_page_header_footer = different
        if not different:
            continue
        # пустая строка = колонтитул первой страницы пустой (например, титульный лист без номера)
        for container, key in ((section.first_page_header, "header_text"),
                               (section.first_page_footer, "footer_text")):
            fill_header_footer(container, {**cmd, "text": cmd.get(key, "")})
    if not different:
        return "одинаковые колонтитулы на всех страницах"
    return "первая страница с отдельными колонтитулами (по умолчанию пустыми)"


# ----- 7. Стили -------------------------------------------------------------
def find_style(doc: Document, name: str):
    wanted = str(name).strip().lower()
    for style in doc.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH and wanted in ((style.name or "").lower(), style.style_id.lower()):
            return style
    return None


def set_style_font(style, spec: dict) -> None:
    font = style.font
    if spec.get("name"):
        font.name = spec["name"]
        rfonts = style.element.rPr.rFonts
        # шрифты темы (asciiTheme и др.) перебивают явное имя, поэтому убираем их
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            rfonts.attrib.pop(qn(attr), None)
        rfonts.set(qn("w:eastAsia"), spec["name"])
        rfonts.set(qn("w:cs"), spec["name"])
    if spec.get("size_pt"):
        font.size = Pt(float(spec["size_pt"]))
    for key in ("bold", "italic", "underline"):
        if spec.get(key) is not None:
            setattr(font, key, bool(spec[key]))
    color = str(spec.get("color", "")).lstrip("#")
    if re.fullmatch(r"[0-9a-fA-F]{6}", color):
        font.color.rgb = RGBColor.from_string(color.upper())


def clear_direct_formatting(paragraphs: list, spec: dict) -> None:
    """
    Убирает «ручное» оформление, которое перебило бы стиль, но только те свойства,
    которые мы меняем в стиле. Жирный/курсив внутри текста не трогаем.
    """
    run_tags = []
    if spec.get("name"):
        run_tags.append("w:rFonts")
    if spec.get("size_pt"):
        run_tags += ["w:sz", "w:szCs"]
    if spec.get("color"):
        run_tags.append("w:color")
    spacing_drop, indent_drop = [], []
    if spec.get("line_spacing") is not None:
        spacing_drop += ["w:line", "w:lineRule"]
    if spec.get("space_before_pt") is not None:
        spacing_drop += ["w:before", "w:beforeLines", "w:beforeAutospacing"]
    if spec.get("space_after_pt") is not None:
        spacing_drop += ["w:after", "w:afterLines", "w:afterAutospacing"]
    if spec.get("first_line_indent_cm") is not None:
        indent_drop += ["w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars"]
    if spec.get("left_indent_cm") is not None:
        indent_drop += ["w:left", "w:start", "w:leftChars", "w:startChars"]

    for p in paragraphs:
        for run in p.runs:
            rPr = run._r.rPr
            if rPr is not None:
                for tag in run_tags:
                    for el in rPr.findall(qn(tag)):
                        rPr.remove(el)
        pPr = p._p.pPr
        if pPr is None:
            continue
        if spec.get("align"):
            for el in pPr.findall(qn("w:jc")):
                pPr.remove(el)
        for tag, drop in (("w:spacing", spacing_drop), ("w:ind", indent_drop)):
            el = pPr.find(qn(tag))
            if el is not None and drop:
                for attr in drop:
                    el.attrib.pop(qn(attr), None)
                if not el.attrib:
                    pPr.remove(el)


def cmd_set_style(doc, cmd):
    style = find_style(doc, cmd.get("style", ""))
    if style is None:
        available = ", ".join(sorted(s.name for s in doc.styles if s.type == WD_STYLE_TYPE.PARAGRAPH))
        raise ValueError(f"стиль «{cmd.get('style')}» не найден. Есть: {available}")

    set_style_font(style, cmd)
    pf = style.paragraph_format
    align = ALIGNMENTS.get(str(cmd.get("align", "")).lower())
    if align is not None:
        pf.alignment = align
    if cmd.get("line_spacing") is not None:
        pf.line_spacing = float(cmd["line_spacing"])
    if cmd.get("space_before_pt") is not None:
        pf.space_before = Pt(float(cmd["space_before_pt"]))
    if cmd.get("space_after_pt") is not None:
        pf.space_after = Pt(float(cmd["space_after_pt"]))
    if cmd.get("first_line_indent_cm") is not None:
        pf.first_line_indent = Cm(float(cmd["first_line_indent_cm"]))
    if cmd.get("left_indent_cm") is not None:
        pf.left_indent = Cm(float(cmd["left_indent_cm"]))

    cleaned = 0
    if cmd.get("clear_direct", True):
        # только основной текст: абзацы в таблицах оформляются командой set_table и не должны сбрасываться
        users = [p for p in doc.paragraphs if p.style.style_id == style.style_id]
        clear_direct_formatting(users, cmd)
        cleaned = len(users)
    return f"стиль «{style.name}» обновлён (ручное оформление снято в {cleaned} абз.)"


# Словарь «название действия -> функция». Новая команда = новая функция + строка тут.
ACTIONS = {
    "set_alignment": cmd_set_alignment,
    "set_font": cmd_set_font,
    "set_paragraph_format": cmd_set_paragraph_format,
    "set_columns": cmd_set_columns,
    "set_header": cmd_set_header,
    "set_footer": cmd_set_footer,
    "set_margins": cmd_set_margins,
    "set_orientation": cmd_set_orientation,
    "set_page_size": cmd_set_page_size,
    "set_list": cmd_set_list,
    "set_table": cmd_set_table,
    "insert_page_break": cmd_insert_page_break,
    "insert_toc": cmd_insert_toc,
    "replace_text": cmd_replace_text,
    "set_first_page": cmd_set_first_page,
    "set_style": cmd_set_style,
}


def apply_commands(doc: Document, commands: list[dict]) -> list[str]:
    log = []
    for cmd in commands:
        action = cmd.get("action")
        handler = ACTIONS.get(action)
        if handler is None:
            log.append(f"⚠️ Неизвестная команда пропущена: {action}")
            continue
        try:
            log.append(f"✅ {action}: {handler(doc, cmd)}")
        except Exception as exc:  # одна плохая команда не ломает остальные
            log.append(f"❌ {action}: ошибка — {exc}")
    return log


# ---------------------------------------------------------------------------
# 3.2 ПРЕСЕТЫ: готовые наборы команд, ИИ не нужен
# ---------------------------------------------------------------------------
ALL = {"type": "all"}
BODY = {"type": "body"}
HEADINGS = {"type": "headings"}

PRESETS = {
    "🎓 ГОСТ (курсовая)": [
        {"action": "set_page_size", "width_cm": 21, "height_cm": 29.7},
        {"action": "set_orientation", "orientation": "portrait"},
        {"action": "set_margins", "top_cm": 2, "bottom_cm": 2, "left_cm": 3, "right_cm": 1.5},
        {"action": "set_columns", "count": 1, "spacing_cm": 1.25, "line_between": False},
        {"action": "set_font", "target": ALL, "name": "Times New Roman", "size_pt": 14, "color": "000000"},
        {"action": "set_alignment", "target": BODY, "align": "justify"},
        {"action": "set_paragraph_format", "target": BODY, "line_spacing": 1.5,
         "space_before_pt": 0, "space_after_pt": 0, "first_line_indent_cm": 1.25},
        {"action": "set_font", "target": HEADINGS, "bold": True},
        {"action": "set_alignment", "target": HEADINGS, "align": "center"},
        {"action": "set_paragraph_format", "target": HEADINGS, "line_spacing": 1.5,
         "space_before_pt": 0, "space_after_pt": 12, "first_line_indent_cm": 0},
        {"action": "set_footer", "text": "{PAGE}", "align": "center",
         "name": "Times New Roman", "size_pt": 12},
    ],
    "📰 Статья (2 колонки)": [
        {"action": "set_page_size", "width_cm": 21, "height_cm": 29.7},
        {"action": "set_margins", "top_cm": 2, "bottom_cm": 2, "left_cm": 2, "right_cm": 2},
        {"action": "set_columns", "count": 2, "spacing_cm": 1.0, "line_between": False},
        {"action": "set_font", "target": ALL, "name": "Times New Roman", "size_pt": 10, "color": "000000"},
        {"action": "set_alignment", "target": BODY, "align": "justify"},
        {"action": "set_paragraph_format", "target": BODY, "line_spacing": 1.0,
         "space_before_pt": 0, "space_after_pt": 0, "first_line_indent_cm": 0.75},
        {"action": "set_font", "target": HEADINGS, "bold": True},
        {"action": "set_footer", "text": "{PAGE}", "align": "center", "size_pt": 9},
    ],
    "💼 Деловой (Arial 11)": [
        {"action": "set_page_size", "width_cm": 21, "height_cm": 29.7},
        {"action": "set_margins", "top_cm": 2, "bottom_cm": 2, "left_cm": 2.5, "right_cm": 1.5},
        {"action": "set_columns", "count": 1, "spacing_cm": 1.25, "line_between": False},
        {"action": "set_font", "target": ALL, "name": "Arial", "size_pt": 11, "color": "000000"},
        {"action": "set_alignment", "target": BODY, "align": "left"},
        {"action": "set_paragraph_format", "target": BODY, "line_spacing": 1.15,
         "space_before_pt": 0, "space_after_pt": 6, "first_line_indent_cm": 0},
        {"action": "set_font", "target": HEADINGS, "bold": True},
        {"action": "set_alignment", "target": HEADINGS, "align": "left"},
        {"action": "set_footer", "text": "Страница {PAGE} из {NUMPAGES}", "align": "right",
         "name": "Arial", "size_pt": 9},
    ],
}

PRESETS["🎓 ГОСТ + титул без номера"] = PRESETS["🎓 ГОСТ (курсовая)"] + [
    {"action": "set_first_page", "different": True, "footer_text": ""},
]

MAX_HISTORY = 20  # сколько версий документа хранить для отмены


# ---------------------------------------------------------------------------
# 4. ИНТЕРФЕЙС STREAMLIT: две страницы — главная и редактор
# Переключение страниц: st.session_state["page"] = "home" | "editor".
# Состояние редактора в session_state:
#   history      - список версий документа (байты), последняя = текущая
#   steps        - подписи версий (для списка «История правок»)
#   pending_json - команды, ожидающие подтверждения (редактируемый текст)
# ---------------------------------------------------------------------------
PROJECT_NAME = "Textio"  # поменяйте на своё название проекта
TAGLINE = "Форматируйте Word словами"

PRESET_INFO = {
    "🎓 ГОСТ (курсовая)": "Times New Roman 14, интервал 1.5, поля 3/2/1.5/2 см, красная строка, номер страницы снизу.",
    "🎓 ГОСТ + титул без номера": "То же, что ГОСТ, но на титульном листе номер страницы не отображается.",
    "📰 Статья (2 колонки)": "Две колонки, Times New Roman 10, выравнивание по ширине, узкие поля.",
    "💼 Деловой (Arial 11)": "Arial 11, интервал 1.15, выравнивание по левому краю, «Страница X из Y» справа.",
}

FEATURES = [
    ("🗣️", "Пишите обычными словами",
     "«Шрифт 14, интервал 1.5, заголовки по центру» — без поиска нужных кнопок в меню Word."),
    ("📋", "Готовые шаблоны",
     "ГОСТ для курсовой, статья в две колонки, деловой документ — оформление в один клик."),
    ("🎯", "Копируйте оформление по образцу",
     "Загрузите красиво оформленный документ: поля, шрифты, колонтитулы и заголовки перенесутся в ваш файл."),
    ("👀", "ИИ под вашим контролем",
     "Перед применением вы видите, что именно понял ИИ, и можете поправить. Любой шаг можно отменить."),
    ("📑", "Сложное оформление",
     "Списки, таблицы, оглавление, разрывы страниц, замена текста, титульный лист без номера."),
    ("🔒", "Документ не уходит в ИИ целиком",
     "Для текстового запроса ИИ получает лишь краткий план: первые слова абзацев. Шаблоны и образец работают без ИИ."),
]

STEPS = [
    ("1", "Загрузите документ", "Подойдёт любой файл в формате .docx."),
    ("2", "Выберите способ", "Текстовый запрос, готовый шаблон или оформление по образцу."),
    ("3", "Проверьте и скачайте", "Посмотрите изменения, отмените лишнее и скачайте готовый файл."),
]

EXAMPLES = [
    "Заголовки по центру, основной текст Times New Roman 14 по ширине",
    "Сделай текст в две колонки и добавь номера страниц внизу",
    "Титульный лист без номера, оглавление в самом начале",
    "Замени «Иванов» на «Петров» во всём документе",
    "Оформи таблицы с рамками и серой шапкой",
]

GLOBAL_CSS = """
<style>
#MainMenu, footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent;}
.block-container {max-width: 1100px; padding-top: 2rem; padding-bottom: 4rem;}
:root {--card-bg: rgba(128,128,128,.08); --card-border: rgba(128,128,128,.28);}

.gradient-text {
  background: linear-gradient(90deg, #6366f1, #a855f7, #ec4899);
  -webkit-background-clip: text; background-clip: text; color: transparent;
}
.hero {text-align: center; padding: 3rem 1rem 1.2rem;}
.badge {
  display: inline-block; padding: .35rem .95rem; border-radius: 999px; margin-bottom: 1.3rem;
  border: 1px solid var(--card-border); background: var(--card-bg); font-size: .85rem;
}
.hero-title {
  font-size: clamp(3.2rem, 10vw, 6rem); font-weight: 800; line-height: 1.05; letter-spacing: -.03em;
  margin-bottom: .5rem;
}
.hero-tagline {font-size: clamp(1.3rem, 3.2vw, 2rem); font-weight: 600; margin-bottom: 1rem;}
.hero-lead {max-width: 680px; margin: 0 auto 1.8rem; font-size: 1.15rem; line-height: 1.65; opacity: .8;}
.hint {text-align: center; opacity: .6; font-size: .9rem; margin-top: .6rem;}

.stats {display: flex; gap: 1rem; justify-content: center; flex-wrap: wrap; margin: 2.2rem 0 1rem;}
.stat {
  min-width: 170px; text-align: center; padding: 1rem 1.4rem; border-radius: 16px;
  border: 1px solid var(--card-border); background: var(--card-bg);
}
.stat-num {font-size: 2rem; font-weight: 800;}
.stat-label {opacity: .7; font-size: .9rem;}

.section-title {text-align: center; font-size: 2rem; font-weight: 700; margin: 3.5rem 0 .4rem;}
.section-sub {text-align: center; opacity: .7; margin-bottom: 1.8rem;}
.cards {display: grid; grid-template-columns: repeat(auto-fit, minmax(290px, 1fr)); gap: 1.1rem;}
.card {
  border: 1px solid var(--card-border); background: var(--card-bg); border-radius: 18px; padding: 1.4rem;
  transition: transform .2s ease, border-color .2s ease;
}
.card:hover {transform: translateY(-4px); border-color: #a855f7;}
.card-icon {font-size: 2rem; margin-bottom: .5rem;}
.card-title {font-size: 1.15rem; font-weight: 700; margin-bottom: .4rem;}
.card-text {opacity: .75; line-height: 1.55; font-size: .96rem;}
.step-num {
  width: 42px; height: 42px; border-radius: 50%; margin-bottom: .8rem; color: #fff; font-weight: 700;
  display: flex; align-items: center; justify-content: center;
  background: linear-gradient(135deg, #6366f1, #ec4899);
}
.chips {display: flex; flex-wrap: wrap; gap: .6rem; justify-content: center;}
.chip {
  padding: .55rem 1rem; border-radius: 999px; font-size: .95rem;
  border: 1px solid var(--card-border); background: var(--card-bg);
}
.cta-text {text-align: center; margin: 3.5rem 0 1.2rem;}
.cta-title {font-size: clamp(1.8rem, 4vw, 2.6rem); font-weight: 800; margin-bottom: .4rem;}
.footer-note {text-align: center; opacity: .5; font-size: .85rem; margin-top: 3rem;}

.editor-title {font-size: 1.6rem; font-weight: 800; padding-top: .3rem;}

/* акцентные кнопки */
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"],
button[data-testid="stBaseButton-primary"] {
  background: linear-gradient(90deg, #6366f1, #a855f7, #ec4899); border: none; color: #fff;
  border-radius: 12px; font-weight: 600; transition: transform .15s ease, filter .15s ease;
}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover {filter: brightness(1.12); transform: translateY(-1px); color: #fff;}
button[data-baseweb="tab"] p {font-size: 1.05rem; font-weight: 600;}
</style>
"""

HOME_CSS = """
<style>
div.stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"] {
  padding: .8rem 1.5rem; border-radius: 14px; box-shadow: 0 10px 30px rgba(168,85,247,.35);
}
div.stButton > button[kind="primary"] p, button[data-testid="stBaseButton-primary"] p {
  font-size: 1.2rem; font-weight: 700;
}
</style>
"""


def html(markup: str) -> None:
    """Выводит HTML. Отступы убираем, иначе Markdown примет строки за блок кода."""
    cleaned = "\n".join(line.strip() for line in markup.splitlines() if line.strip())
    st.markdown(cleaned, unsafe_allow_html=True)


# --- состояние и обработчики кнопок (callbacks) ---
def init_state():
    st.session_state.setdefault("page", "home")
    st.session_state.setdefault("history", [])
    st.session_state.setdefault("steps", [])
    st.session_state.setdefault("file_key", None)
    st.session_state.setdefault("doc_name", "document.docx")
    st.session_state.setdefault("pending_label", "")


def go(page: str):
    st.session_state["page"] = page


def current_doc() -> Document:
    return Document(io.BytesIO(st.session_state["history"][-1]))


def push_version(doc: Document, label: str) -> None:
    buffer = io.BytesIO()
    doc.save(buffer)
    st.session_state["history"].append(buffer.getvalue())
    st.session_state["steps"].append(label)
    if len(st.session_state["history"]) > MAX_HISTORY:  # первую (исходную) версию не трогаем
        del st.session_state["history"][1]
        del st.session_state["steps"][1]


def run_commands(commands: list[dict], label: str) -> None:
    doc = current_doc()
    log = apply_commands(doc, commands)
    push_version(doc, label)
    st.session_state["flash"] = ("success", f"Применено: {label}", log)


def on_preset(name: str):
    run_commands(PRESETS[name], name)


def on_apply_pending():
    try:
        commands = parse_commands(st.session_state.get("pending_json", ""))
    except Exception as exc:
        st.session_state["flash"] = ("error", f"Не удалось разобрать JSON: {exc}", [])
        return
    if not commands:
        st.session_state["flash"] = ("warning", "В JSON нет ни одной команды.", [])
        return
    run_commands(commands, st.session_state.get("pending_label") or "правка")
    st.session_state.pop("pending_json", None)


def on_discard_pending():
    st.session_state.pop("pending_json", None)


def on_undo():
    if len(st.session_state["history"]) > 1:
        st.session_state["history"].pop()
        removed = st.session_state["steps"].pop()
        st.session_state["flash"] = ("info", f"Отменено: {removed}", [])


def on_reset():
    del st.session_state["history"][1:]
    del st.session_state["steps"][1:]
    st.session_state.pop("pending_json", None)
    st.session_state["flash"] = ("info", "Документ возвращён к исходному виду.", [])


# --- формирование команд (результат попадает на проверку, а не применяется сразу) ---
def set_pending(commands: list[dict], label: str) -> None:
    if not commands:
        st.warning("Не нашлось подходящих команд. Попробуйте переформулировать запрос.")
        return
    st.session_state["pending_json"] = json.dumps({"commands": commands}, ensure_ascii=False, indent=2)
    st.session_state["pending_label"] = label
    st.toast("Команды готовы: проверьте их ниже и нажмите «Применить»")


def generate_from_prompt(prompt: str, api_key: str | None) -> None:
    if not prompt.strip():
        st.warning("Напишите, что нужно изменить в документе.")
        return
    if not api_key:
        st.error("Не найден API-ключ Gemini. Введите его выше или задайте переменную GEMINI_API_KEY.")
        return
    try:
        with st.spinner("Gemini формирует команды (при перегрузке может занять до минуты)..."):
            commands, _model = ask_gemini(api_key, prompt, build_outline(current_doc()))
    except Exception as exc:
        st.error(f"Ошибка при обращении к Gemini: {exc}")
        return
    set_pending(commands, f"запрос «{prompt.strip()[:40]}»")


def generate_from_sample(example_file, copy_hf_text: bool) -> None:
    if example_file is None:
        st.warning("Сначала загрузите документ-образец.")
        return
    try:
        example_doc = Document(io.BytesIO(example_file.getvalue()))
        commands = extract_format_commands(example_doc, current_doc(), copy_hf_text)
    except Exception as exc:
        st.error(f"Не удалось прочитать образец: {exc}")
        return
    set_pending(commands, "образец")


# ---------------------------------------------------------------------------
# СТРАНИЦА 1: ГЛАВНАЯ
# ---------------------------------------------------------------------------
def render_home() -> None:
    html(HOME_CSS)
    html(f"""
    <div class="hero">
      <div class="badge">✨ ИИ-помощник для документов Word</div>
      <div class="hero-title"><span class="gradient-text">{PROJECT_NAME}</span></div>
      <div class="hero-tagline">{TAGLINE}</div>
      <div class="hero-lead">
        Загрузите документ и опишите, что нужно изменить, выберите готовый шаблон
        или скопируйте оформление с образца. Шрифты, поля, колонтитулы, оглавление
        и таблицы будут оформлены без возни с меню Word.
      </div>
    </div>
    """)
    _, middle, _ = st.columns([1.2, 1, 1.2])
    with middle:
        st.button("🚀 Начать прямо сейчас", type="primary", on_click=go, args=("editor",),
                  use_container_width=True, key="cta_top")
    html('<div class="hint">Без регистрации · шаблоны и образец работают без ИИ-ключа</div>')

    html("""
    <div class="stats">
      <div class="stat"><div class="stat-num gradient-text">3</div><div class="stat-label">способа оформления</div></div>
      <div class="stat"><div class="stat-num gradient-text">16</div><div class="stat-label">операций над документом</div></div>
      <div class="stat"><div class="stat-num gradient-text">20</div><div class="stat-label">шагов отмены</div></div>
    </div>
    """)

    html('<div class="section-title">Почему это удобно</div>'
         '<div class="section-sub">Меньше рутины — больше времени на сам текст</div>')
    cards = "".join(
        f'<div class="card"><div class="card-icon">{icon}</div>'
        f'<div class="card-title">{title}</div><div class="card-text">{text}</div></div>'
        for icon, title, text in FEATURES
    )
    html(f'<div class="cards">{cards}</div>')

    html('<div class="section-title">Как это работает</div>'
         '<div class="section-sub">Три шага от исходного файла до готового</div>')
    steps = "".join(
        f'<div class="card"><div class="step-num">{num}</div>'
        f'<div class="card-title">{title}</div><div class="card-text">{text}</div></div>'
        for num, title, text in STEPS
    )
    html(f'<div class="cards">{steps}</div>')

    html('<div class="section-title">Что можно попросить</div>'
         '<div class="section-sub">Примеры запросов — формулируйте своими словами</div>')
    chips = "".join(f'<div class="chip">«{example}»</div>' for example in EXAMPLES)
    html(f'<div class="chips">{chips}</div>')

    html('<div class="cta-text"><div class="cta-title">Готовы оформить документ?</div>'
         '<div class="section-sub">Это займёт пару минут</div></div>')
    _, middle, _ = st.columns([1.2, 1, 1.2])
    with middle:
        st.button("🚀 Начать прямо сейчас", type="primary", on_click=go, args=("editor",),
                  use_container_width=True, key="cta_bottom")
    html(f'<div class="footer-note">{PROJECT_NAME} · Streamlit · Gemini · python-docx</div>')


# ---------------------------------------------------------------------------
# СТРАНИЦА 2: РЕДАКТОР
# ---------------------------------------------------------------------------
def render_editor() -> None:
    back_col, title_col = st.columns([1, 4])
    back_col.button("← На главную", on_click=go, args=("home",), key="back_home")
    title_col.markdown(f'<div class="editor-title">📝 Редактор · <span class="gradient-text">{PROJECT_NAME}</span></div>',
                       unsafe_allow_html=True)

    # --- загрузка документа ---
    has_doc = bool(st.session_state["history"])
    uploaded = st.file_uploader(
        "Заменить документ" if has_doc else "Загрузите файл Word (.docx), который нужно оформить",
        type=["docx"], key="target",
    )
    if uploaded is not None:
        file_key = f"{uploaded.name}:{uploaded.size}"
        if st.session_state["file_key"] != file_key:  # новый файл -> история начинается заново
            try:
                Document(io.BytesIO(uploaded.getvalue()))  # проверка, что файл открывается
            except Exception as exc:
                st.error(f"Не удалось открыть файл (он точно .docx?): {exc}")
                st.stop()
            st.session_state["history"] = [uploaded.getvalue()]
            st.session_state["steps"] = ["Исходный файл"]
            st.session_state["file_key"] = file_key
            st.session_state["doc_name"] = uploaded.name
            st.session_state.pop("pending_json", None)
            st.session_state.pop("flash", None)
    if not st.session_state["history"]:
        st.info("Загрузите .docx-файл, чтобы начать. Дальше можно оформить его запросом, шаблоном или по образцу.")
        st.stop()

    # --- сообщение о результате последнего действия ---
    flash = st.session_state.pop("flash", None)
    if flash:
        kind, text, log = flash
        getattr(st, kind)(text)
        if log:
            with st.expander("Что было сделано"):
                for line in log:
                    st.write(line)

    left, right = st.columns([3, 2], gap="large")

    with left:
        tab_text, tab_preset, tab_sample = st.tabs(["✍️ Текстовый запрос", "📋 По шаблону", "🎯 По образцу"])

        with tab_text:
            st.caption("Опишите словами, что изменить. ИИ переведёт это в команды, а вы проверите их перед применением.")
            typed_key = st.text_input(
                "Gemini API-ключ", type="password", key="api_key_input",
                help="Бесплатный ключ создаётся на aistudio.google.com. Можно задать переменную GEMINI_API_KEY "
                     "или secrets.toml, тогда вводить не нужно.",
            )
            user_prompt = st.text_area(
                "Что нужно сделать с документом?", key="prompt_text", height=130,
                placeholder="Например: заголовки по центру, шрифт Arial 12, интервал 1.5, "
                            "внизу номер страницы, титульный лист без номера",
            )
            if st.button("Сформировать команды", type="primary", key="gen_text"):
                generate_from_prompt(user_prompt, typed_key or get_api_key())
            st.caption(f"Модель: `{MODEL_NAME}`")

        with tab_preset:
            st.caption("Готовые наборы оформления. Применяются сразу, ИИ и ключ не нужны.")
            for name in PRESETS:
                with st.container(border=True):
                    st.markdown(f"**{name}**")
                    st.caption(PRESET_INFO.get(name, ""))
                    st.button("Применить", key=f"preset_{name}", on_click=on_preset, args=(name,))

        with tab_sample:
            st.caption("Загрузите документ с нужным оформлением: размер страницы, поля, колонки, колонтитулы, "
                       "шрифты, интервалы и заголовки будут перенесены в ваш файл. ИИ и ключ не нужны.")
            example_file = st.file_uploader("Документ-образец (.docx)", type=["docx"], key="example")
            copy_hf_text = st.checkbox(
                "Копировать текст колонтитулов из образца", value=True,
                help="Если выключить, скопируются только колонтитулы с номерами страниц, "
                     "а обычный текст (фамилия, название работы) нет.",
            )
            if st.button("Сформировать команды по образцу", type="primary", key="gen_sample"):
                generate_from_sample(example_file, copy_hf_text)

        # --- проверка команд перед применением (общая для вкладок «запрос» и «образец») ---
        if "pending_json" in st.session_state:
            st.divider()
            st.markdown("#### 🔍 Проверьте команды")
            st.caption("Это то, что будет применено к документу. JSON можно отредактировать вручную.")
            st.text_area("JSON-команды", key="pending_json", height=320, label_visibility="collapsed")
            apply_col, discard_col = st.columns(2)
            apply_col.button("✅ Применить", on_click=on_apply_pending, type="primary",
                             use_container_width=True, key="apply_pending")
            discard_col.button("🗑 Отклонить", on_click=on_discard_pending,
                               use_container_width=True, key="discard_pending")

    with right:
        with st.container(border=True):
            st.markdown("#### 📄 Результат")
            st.caption(f"{st.session_state['doc_name']} · версия {len(st.session_state['history'])}")
            st.download_button(
                "⬇️ Скачать документ",
                data=st.session_state["history"][-1],
                file_name=f"formatted_{st.session_state['doc_name']}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary", use_container_width=True,
            )
            undo_col, reset_col = st.columns(2)
            can_undo = len(st.session_state["history"]) > 1
            undo_col.button("↩️ Отменить", on_click=on_undo, disabled=not can_undo,
                            use_container_width=True, key="undo")
            reset_col.button("⏮ Сбросить", on_click=on_reset, disabled=not can_undo,
                             use_container_width=True, key="reset")
        with st.expander(f"История правок ({len(st.session_state['steps'])})"):
            for i, step in enumerate(st.session_state["steps"]):
                st.write(f"{i}. {step}")


# ---------------------------------------------------------------------------
# ТОЧКА ВХОДА
# ---------------------------------------------------------------------------
st.set_page_config(page_title=f"{PROJECT_NAME}: форматирование Word с ИИ", page_icon="📝",
                   layout="wide", initial_sidebar_state="collapsed")
init_state()
html(GLOBAL_CSS)
if st.session_state["page"] == "editor":
    render_editor()
else:
    render_home()
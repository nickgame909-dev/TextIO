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

import streamlit as st
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

"target" object (used by set_alignment, set_font, set_paragraph_format):
  {"type": "all" | "headings" | "body" | "paragraphs" | "contains",
   "indexes": [1, 2, -1],      // only for "paragraphs": 1-based, negative = from the end
   "text": "substring"}        // only for "contains"
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
    return "\n".join(lines) or "(документ пустой)"


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


def add_field(paragraph, instruction: str):
    """Вставляет поле Word (PAGE, NUMPAGES) — номер страницы обновляется сам."""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "1"
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
# 4. ИНТЕРФЕЙС STREAMLIT
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Word-форматтер на ИИ", page_icon="📝")
st.title("📝 Форматирование Word с помощью ИИ")
st.caption("Загрузите .docx, опишите словами, что изменить, — и скачайте результат.")

with st.sidebar:
    st.header("Настройки")
    typed_key = st.text_input("Gemini API-ключ", type="password",
                              help="Необязательно, если ключ задан в переменной окружения GEMINI_API_KEY")
    st.markdown(f"Модель: `{MODEL_NAME}`")

api_key = typed_key or get_api_key()

uploaded = st.file_uploader("1. Загрузите файл Word", type=["docx"])
user_prompt = st.text_area(
    "2. Что нужно сделать с документом?",
    placeholder="Например: сделай текст в две колонки, шрифт Times New Roman 14, "
                "выравнивание по ширине, заголовки по центру, внизу номер страницы",
    height=120,
)

if st.button("Применить форматирование", type="primary"):
    if not api_key:
        st.error("Не найден API-ключ Gemini. Введите его в боковой панели или задайте GEMINI_API_KEY.")
    elif uploaded is None:
        st.warning("Сначала загрузите .docx-файл.")
    elif not user_prompt.strip():
        st.warning("Напишите, что нужно изменить.")
    else:
        try:
            doc = Document(io.BytesIO(uploaded.getvalue()))
        except Exception as exc:
            st.error(f"Не удалось открыть файл (он точно .docx?): {exc}")
            st.stop()

        try:
            with st.spinner("Gemini формирует команды (при перегрузке сервера может занять до минуты)..."):
                commands, used_model = ask_gemini(api_key, user_prompt, build_outline(doc))
        except Exception as exc:
            st.error(f"Ошибка при обращении к Gemini: {exc}")
            st.stop()

        if not commands:
            st.warning("Модель не нашла подходящих команд. Попробуйте переформулировать запрос.")
            st.session_state.pop("result", None)
        else:
            log = apply_commands(doc, commands)
            buffer = io.BytesIO()
            doc.save(buffer)
            st.session_state["result"] = {
                "data": buffer.getvalue(),
                "name": f"formatted_{uploaded.name}",
                "commands": commands,
                "model": used_model,
                "log": log,
            }

# Результат хранится в session_state, чтобы кнопка скачивания не пропадала после перезапуска страницы
result = st.session_state.get("result")
if result:
    st.success("Готово! Документ обработан.")
    st.caption(f"Ответила модель: `{result['model']}`")
    st.download_button(
        "⬇️ Скачать изменённый файл",
        data=result["data"],
        file_name=result["name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    with st.expander("Что было сделано"):
        for line in result["log"]:
            st.write(line)
    with st.expander("JSON-команды от Gemini"):
        st.json(result["commands"])

"""Live Word tools. The agent drives the REAL Word window through COM (pywin32).

Windows + desktop Word required. Paragraph numbers are 1-based, as in Word.
"""

import os
import time

from com_worker import ComWorker, dispatch

STEP_DELAY = float(os.getenv("OFFICE_STEP_DELAY", os.getenv("EXCEL_STEP_DELAY", "0.4")))

_worker = ComWorker()
_word = None  # only touched from the worker thread

# Word built-in style ids (work in any Word language)
_STYLE_IDS = {
    "normal": -1,
    "title": -63,
    "subtitle": -75,
    "quote": -181,
    "no spacing": -158,
    "list bullet": -49,
    "list number": -50,
}
for _n in range(1, 10):
    _STYLE_IDS[f"heading {_n}"] = -(_n + 1)

_ALIGN = {"left": 0, "center": 1, "right": 2, "justify": 3}
_HIGHLIGHT = {"none": 0, "yellow": 7, "green": 4, "turquoise": 3, "pink": 5}


# --------------------------------------------------------------------------
# Helpers (run on the worker thread)
# --------------------------------------------------------------------------
def _app():
    global _word
    if _word is not None:
        try:
            _word.Documents.Count
        except Exception:  # user closed Word
            _word = None
    if _word is None:
        import win32com.client

        _word = win32com.client.Dispatch("Word.Application")
        _word.Visible = True
    return _word


def _doc():
    app = _app()
    if app.Documents.Count == 0:
        app.Documents.Add()
    return app.ActiveDocument


def _para(doc, index):
    n = doc.Paragraphs.Count
    index = int(index)
    if not 1 <= index <= n:
        raise ValueError(f"Paragraph {index} does not exist. The document has {n} paragraphs.")
    return doc.Paragraphs.Item(index)


def _bgr(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    return r + g * 256 + b * 65536


def _style_name(p) -> str:
    try:
        return p.Style.NameLocal
    except Exception:
        return str(p.Style)


def _set_style(para, style):
    key = str(style or "Normal").strip()
    value = _STYLE_IDS.get(key.lower(), key)
    try:
        para.Style = value
    except Exception:
        para.Style = key


def _show(rng=None):
    """Scroll to and select what was just changed so the user can follow along."""
    try:
        (rng or _doc().Content).Select()
        _app().ScreenRefresh()
    except Exception:
        pass
    time.sleep(STEP_DELAY)


def _blank_last_paragraph(doc):
    """Return an empty paragraph at the end of the document (adding one if needed)."""
    last = doc.Paragraphs.Item(doc.Paragraphs.Count)
    if last.Range.Text.strip("\r\n\x07\x0c ") != "":
        doc.Content.InsertParagraphAfter()
        last = doc.Paragraphs.Item(doc.Paragraphs.Count)
    return last


def _style_range(rng, bold=None, italic=None, underline=None, font_size=None,
                 font_name=None, font_color=None, highlight=None):
    f = rng.Font
    if bold is not None:
        f.Bold = bool(bold)
    if italic is not None:
        f.Italic = bool(italic)
    if underline is not None:
        f.Underline = 1 if underline else 0
    if font_size:
        f.Size = font_size
    if font_name:
        f.Name = font_name
    if font_color:
        f.Color = _bgr(font_color)
    if highlight:
        rng.HighlightColorIndex = _HIGHLIGHT.get(highlight, 0)


def _add_one_paragraph(doc, text, style="Normal", align=None, **fmt):
    last = _blank_last_paragraph(doc)
    _set_style(last, style)
    try:
        last.Reset()  # drop leftover manual formatting inherited from the previous paragraph
        last.Range.Font.Reset()
    except Exception:
        pass
    last.Range.InsertBefore(text)
    last = doc.Paragraphs.Item(doc.Paragraphs.Count)
    if align:
        last.Alignment = _ALIGN[align]
    if any(v is not None for v in fmt.values()):
        _style_range(last.Range, **fmt)
    return last


# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------
def open_document(path=None):
    app = _app()
    if path:
        path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        doc = app.Documents.Open(path)
    else:
        doc = app.Documents.Add()
    app.Activate()
    return {"document": doc.Name, "paragraphs": doc.Paragraphs.Count}


def get_document_info(preview_paragraphs=60):
    doc = _doc()
    n = doc.Paragraphs.Count
    lines = []
    for i in range(1, min(n, int(preview_paragraphs)) + 1):
        p = doc.Paragraphs.Item(i)
        text = p.Range.Text.strip("\r\n\x07\x0c ")
        lines.append(f"{i} [{_style_name(p)}] {text[:100]}")
    return {
        "document": doc.Name,
        "path": doc.FullName,
        "paragraph_count": n,
        "table_count": doc.Tables.Count,
        "page_count": doc.ComputeStatistics(2),
        "word_count": doc.ComputeStatistics(0),
        "paragraphs_preview": lines,
    }


def read_document(start_paragraph=1, end_paragraph=None, max_chars=6000):
    doc = _doc()
    n = doc.Paragraphs.Count
    start = max(1, int(start_paragraph))
    end = min(n, int(end_paragraph)) if end_paragraph else n
    out, total = [], 0
    for i in range(start, end + 1):
        p = doc.Paragraphs.Item(i)
        line = f"{i} [{_style_name(p)}] {p.Range.Text.strip(chr(13) + chr(7) + chr(12))}"
        total += len(line)
        if total > max_chars:
            out.append(f"... stopped at paragraph {i - 1}; call again with start_paragraph={i}")
            break
        out.append(line)
    return {"paragraph_count": n, "paragraphs": out}


def add_paragraph(text, style="Normal", align=None, bold=None, italic=None, underline=None,
                  font_size=None, font_name=None, font_color=None):
    """Append paragraph(s) at the end. A '\\n' in text starts a new paragraph."""
    doc = _doc()
    count = 0
    for line in str(text).split("\n"):
        if not line.strip():
            continue
        _add_one_paragraph(doc, line, style, align, bold=bold, italic=italic, underline=underline,
                           font_size=font_size, font_name=font_name, font_color=font_color)
        count += 1
        time.sleep(0.1)
    _show()
    return {"paragraphs_added": count, "total_paragraphs": doc.Paragraphs.Count}


def add_heading(text, level=1):
    doc = _doc()
    _add_one_paragraph(doc, text, f"Heading {int(level)}")
    _show()
    return {"added_heading": text, "level": level, "total_paragraphs": doc.Paragraphs.Count}


def add_list(items, ordered=False):
    doc = _doc()
    style = "List Number" if ordered else "List Bullet"
    for item in items:
        _add_one_paragraph(doc, str(item), style)
        time.sleep(0.1)
    _show()
    return {"items_added": len(items), "total_paragraphs": doc.Paragraphs.Count}


def add_table(values, has_header=True, header_fill="#D9E2F3"):
    doc = _doc()
    if not values or not isinstance(values[0], list):
        raise ValueError("values must be a list of rows, e.g. [['Name','Marks'],['Asha',91]]")
    nrows = len(values)
    ncols = max(len(r) for r in values)
    if nrows * ncols > 2000:
        raise ValueError("Table too large (max 2000 cells).")
    last = _blank_last_paragraph(doc)
    _set_style(last, "Normal")
    table = doc.Tables.Add(last.Range, nrows, ncols)
    try:
        table.Style = "Table Grid"
    except Exception:
        table.Borders.Enable = True
    for r, row in enumerate(values, start=1):
        for c, v in enumerate(row, start=1):
            table.Cell(r, c).Range.Text = "" if v is None else str(v)
    if has_header:
        hdr = table.Rows.Item(1)
        hdr.Range.Font.Bold = True
        if header_fill:
            hdr.Shading.BackgroundPatternColor = _bgr(header_fill)
    table.AutoFitBehavior(2)  # fit to window
    _show(table.Range)
    return {"table_added": f"{nrows} rows x {ncols} columns", "tables_in_document": doc.Tables.Count}


def insert_image(path, width_inches=None):
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path):
        raise FileNotFoundError(f"Image not found: {path}")
    doc = _doc()
    last = _blank_last_paragraph(doc)
    _set_style(last, "Normal")
    shape = doc.InlineShapes.AddPicture(path, False, True, last.Range)
    if width_inches:
        shape.LockAspectRatio = True
        shape.Width = float(width_inches) * 72
    _show()
    return {"image_inserted": os.path.basename(path)}


def format_text(find_text=None, paragraph=None, bold=None, italic=None, underline=None,
                font_size=None, font_name=None, font_color=None, highlight=None,
                align=None, match_case=False):
    """Format every occurrence of find_text, or a whole paragraph by number."""
    doc = _doc()
    fmt = dict(bold=bold, italic=italic, underline=underline, font_size=font_size,
               font_name=font_name, font_color=font_color, highlight=highlight)
    if paragraph:
        p = _para(doc, paragraph)
        _style_range(p.Range, **fmt)
        if align:
            p.Alignment = _ALIGN[align]
        _show(p.Range)
        return {"formatted_paragraph": int(paragraph)}
    if not find_text:
        raise ValueError("Give either find_text or paragraph.")
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    f.Text = find_text
    f.Forward = True
    f.Wrap = 0
    f.MatchCase = bool(match_case)
    f.MatchWildcards = False
    count = 0
    while f.Execute() and count < 1000:
        _style_range(rng, **fmt)
        if align:
            rng.ParagraphFormat.Alignment = _ALIGN[align]
        count += 1
        rng.Collapse(0)  # wdCollapseEnd: continue after this match
    _show()
    return {"occurrences_formatted": count}


def replace_text(find_text, replace_with, match_case=False):
    doc = _doc()
    haystack = doc.Content.Text
    needle = find_text
    if not match_case:
        haystack, needle = haystack.lower(), needle.lower()
    count = haystack.count(needle)
    if count:
        f = doc.Content.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        # Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        #         MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace)
        f.Execute(find_text, bool(match_case), False, False, False, False, True, 1, False,
                  replace_with, 2)
    _show()
    return {"replaced_occurrences": count}


def edit_paragraph(paragraph, new_text):
    """Replace the text of one paragraph, keeping its style."""
    doc = _doc()
    p = _para(doc, paragraph)
    rng = p.Range
    rng.MoveEnd(1, -1)  # exclude the paragraph mark
    rng.Text = new_text
    _show(doc.Paragraphs.Item(int(paragraph)).Range)
    return {"edited_paragraph": int(paragraph)}


def insert_paragraph_after(paragraph, text, style="Normal"):
    doc = _doc()
    p = _para(doc, paragraph)
    p.Range.InsertParagraphAfter()
    new_p = doc.Paragraphs.Item(int(paragraph) + 1)
    _set_style(new_p, style)
    try:
        new_p.Reset()
        new_p.Range.Font.Reset()
    except Exception:
        pass
    new_p.Range.InsertBefore(text)
    _show(doc.Paragraphs.Item(int(paragraph) + 1).Range)
    return {"inserted_as_paragraph": int(paragraph) + 1, "total_paragraphs": doc.Paragraphs.Count}


def delete_paragraph(paragraph):
    doc = _doc()
    p = _para(doc, paragraph)
    p.Range.Delete()
    _show()
    return {"deleted_paragraph": int(paragraph), "total_paragraphs": doc.Paragraphs.Count}


def add_page_break():
    doc = _doc()
    rng = doc.Content
    rng.Collapse(0)
    rng.InsertBreak(7)  # wdPageBreak
    _show()
    return {"page_break_added": True}


def set_header_footer(header_text=None, footer_text=None, page_numbers=False):
    doc = _doc()
    sec = doc.Sections.Item(1)
    if header_text is not None:
        sec.Headers.Item(1).Range.Text = header_text
    if footer_text is not None:
        sec.Footers.Item(1).Range.Text = footer_text
    if page_numbers:
        sec.Footers.Item(1).PageNumbers.Add(1, True)  # centered
    _show()
    return {"header_footer_updated": True}


def set_page_setup(orientation=None, margin_inches=None):
    doc = _doc()
    ps = doc.PageSetup
    if orientation:
        ps.Orientation = 1 if orientation == "landscape" else 0
    if margin_inches:
        m = float(margin_inches) * 72
        ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = m
    _show()
    return {"page_setup_updated": True}


def save_document(path=None, overwrite=False):
    """Save only when the user asks. Never overwrites an existing file unless overwrite=True."""
    app = _app()
    doc = _doc()
    if path:
        if not path.lower().endswith((".docx", ".doc")):
            path += ".docx"
        if not os.path.isabs(path):
            path = os.path.join(os.path.expanduser("~"), "Documents", path)
        if os.path.exists(path) and not overwrite:
            raise FileExistsError(f"{path} already exists. Ask the user before overwriting.")
        app.DisplayAlerts = 0
        try:
            doc.SaveAs2(path, 16)  # wdFormatXMLDocument (.docx)
        finally:
            app.DisplayAlerts = -1
        return {"saved_to": path}
    if not doc.Path:
        raise ValueError("This document has never been saved. Provide a file name in 'path'.")
    doc.Save()
    return {"saved_to": doc.FullName}


def export_pdf(path, overwrite=False):
    doc = _doc()
    if not path.lower().endswith(".pdf"):
        path += ".pdf"
    if not os.path.isabs(path):
        path = os.path.join(os.path.expanduser("~"), "Documents", path)
    if os.path.exists(path) and not overwrite:
        raise FileExistsError(f"{path} already exists. Ask the user before overwriting.")
    doc.ExportAsFixedFormat(path, 17)  # wdExportFormatPDF
    return {"pdf_saved_to": path}


TOOLS = {
    f.__name__: f
    for f in [
        open_document, get_document_info, read_document, add_paragraph, add_heading, add_list,
        add_table, insert_image, format_text, replace_text, edit_paragraph,
        insert_paragraph_after, delete_paragraph, add_page_break, set_header_footer,
        set_page_setup, save_document, export_pdf,
    ]
}


# --------------------------------------------------------------------------
# Schemas the LLM sees
# --------------------------------------------------------------------------
def _schema(name, description, properties, required=()):
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": description,
            "parameters": {"type": "object", "properties": properties, "required": list(required)},
        },
    }


_PARA = {"type": "integer", "description": "Paragraph number (1-based), as shown by get_document_info."}
_COLOR = {"type": "string", "description": "Hex color like '#1F4E78'."}
_ALIGN_PROP = {"type": "string", "enum": ["left", "center", "right", "justify"]}

TOOL_SCHEMAS = [
    _schema("open_document",
            "Open the real Word window. With 'path' opens that existing .docx; without it creates a new blank document.",
            {"path": {"type": "string", "description": "Full path of an existing Word file (optional)."}}),
    _schema("get_document_info",
            "See the open document: name, page/word/table counts and a numbered preview of paragraphs with their styles. Call before editing an existing document.",
            {"preview_paragraphs": {"type": "integer"}}),
    _schema("read_document",
            "Read paragraphs (numbered, with styles) from start_paragraph to end_paragraph.",
            {"start_paragraph": {"type": "integer"}, "end_paragraph": {"type": "integer"}}),
    _schema("add_paragraph",
            "Append paragraph(s) at the END of the document. Use style for structure ('Normal', 'Title', 'Subtitle', 'Quote', 'Heading 1'..'Heading 9'). A newline in text starts a new paragraph.",
            {"text": {"type": "string"}, "style": {"type": "string"}, "align": _ALIGN_PROP,
             "bold": {"type": "boolean"}, "italic": {"type": "boolean"}, "underline": {"type": "boolean"},
             "font_size": {"type": "number"}, "font_name": {"type": "string"}, "font_color": _COLOR},
            ["text"]),
    _schema("add_heading", "Append a heading at the end of the document (level 1-9).",
            {"text": {"type": "string"}, "level": {"type": "integer"}}, ["text"]),
    _schema("add_list", "Append a bulleted or numbered list at the end of the document.",
            {"items": {"type": "array", "items": {"type": "string"}}, "ordered": {"type": "boolean"}},
            ["items"]),
    _schema("add_table",
            "Append a table at the end. values is a list of rows; the first row is the header when has_header is true.",
            {"values": {"type": "array", "items": {"type": "array", "items": {"type": ["string", "number", "boolean", "null"]}}},
             "has_header": {"type": "boolean"}, "header_fill": _COLOR},
            ["values"]),
    _schema("insert_image", "Append an image from a file path at the end of the document.",
            {"path": {"type": "string"}, "width_inches": {"type": "number"}}, ["path"]),
    _schema("format_text",
            "Format every occurrence of find_text, OR a whole paragraph by number. Set only what should change.",
            {"find_text": {"type": "string"}, "paragraph": _PARA, "bold": {"type": "boolean"},
             "italic": {"type": "boolean"}, "underline": {"type": "boolean"}, "font_size": {"type": "number"},
             "font_name": {"type": "string"}, "font_color": _COLOR,
             "highlight": {"type": "string", "enum": list(_HIGHLIGHT)}, "align": _ALIGN_PROP,
             "match_case": {"type": "boolean"}}),
    _schema("replace_text", "Find and replace text everywhere in the document.",
            {"find_text": {"type": "string"}, "replace_with": {"type": "string"}, "match_case": {"type": "boolean"}},
            ["find_text", "replace_with"]),
    _schema("edit_paragraph", "Replace the text of one existing paragraph, keeping its style.",
            {"paragraph": _PARA, "new_text": {"type": "string"}}, ["paragraph", "new_text"]),
    _schema("insert_paragraph_after", "Insert a new paragraph right after an existing paragraph number.",
            {"paragraph": _PARA, "text": {"type": "string"}, "style": {"type": "string"}},
            ["paragraph", "text"]),
    _schema("delete_paragraph", "Delete one paragraph by number.", {"paragraph": _PARA}, ["paragraph"]),
    _schema("add_page_break", "Insert a page break at the end of the document.", {}),
    _schema("set_header_footer", "Set header text, footer text and/or centered page numbers.",
            {"header_text": {"type": "string"}, "footer_text": {"type": "string"}, "page_numbers": {"type": "boolean"}}),
    _schema("set_page_setup", "Set page orientation and/or equal margins on all sides.",
            {"orientation": {"type": "string", "enum": ["portrait", "landscape"]}, "margin_inches": {"type": "number"}}),
    _schema("save_document",
            "Save the document. ONLY call when the user asks to save. Never overwrite an existing file unless the user explicitly agrees. A relative name is saved in the Documents folder.",
            {"path": {"type": "string"}, "overwrite": {"type": "boolean"}}),
    _schema("export_pdf", "Export the document as a PDF. Only when the user asks.",
            {"path": {"type": "string"}, "overwrite": {"type": "boolean"}}, ["path"]),
]


def call_tool(name: str, args: dict) -> str:
    return dispatch(_worker, TOOLS, name, args)
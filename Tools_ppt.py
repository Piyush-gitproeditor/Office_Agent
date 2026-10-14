"""Live PowerPoint tools. The agent drives the REAL PowerPoint window through COM (pywin32).

Windows + desktop PowerPoint required. Slide numbers are 1-based. Positions and sizes are in
points (a standard 16:9 slide is 960 x 540 points).
"""

import os
import re
import time

from com_worker import ComWorker, dispatch

STEP_DELAY = float(os.getenv("OFFICE_STEP_DELAY", os.getenv("EXCEL_STEP_DELAY", "0.4")))

_worker = ComWorker()
_ppt = None  # only touched from the worker thread

# Legacy PpSlideLayout ids: work in any PowerPoint language and map to the deck's own layouts
_LAYOUTS = {
    "title": 1,
    "title_and_content": 2,
    "two_content": 3,
    "title_only": 11,
    "blank": 12,
    "section_header": 33,
}
_ALIGN = {"left": 1, "center": 2, "right": 3, "justify": 4}
_SHAPE_TYPES = {1: "autoshape", 3: "chart", 6: "group", 7: "ole_object", 13: "picture",
                14: "placeholder", 17: "textbox", 19: "table", 21: "media"}


# --------------------------------------------------------------------------
# Helpers (run on the worker thread)
# --------------------------------------------------------------------------
def _app():
    global _ppt
    if _ppt is not None:
        try:
            _ppt.Presentations.Count
        except Exception:  # user closed PowerPoint
            _ppt = None
    if _ppt is None:
        import win32com.client

        _ppt = win32com.client.Dispatch("PowerPoint.Application")
        _ppt.Visible = True
    return _ppt


def _pres():
    app = _app()
    if app.Presentations.Count == 0:
        app.Presentations.Add()
    return app.ActivePresentation


def _slide(pres, n):
    n = int(n)
    count = pres.Slides.Count
    if not 1 <= n <= count:
        raise ValueError(f"Slide {n} does not exist. The presentation has {count} slides.")
    return pres.Slides.Item(n)


def _shape(slide, ref):
    shapes = slide.Shapes
    if isinstance(ref, int) or str(ref).isdigit():
        i = int(ref)
        if not 1 <= i <= shapes.Count:
            raise ValueError(f"Shape {i} does not exist on this slide ({shapes.Count} shapes).")
        return shapes.Item(i)
    for i in range(1, shapes.Count + 1):
        if shapes.Item(i).Name.lower() == str(ref).lower():
            return shapes.Item(i)
    names = [shapes.Item(i).Name for i in range(1, shapes.Count + 1)]
    raise ValueError(f"Shape '{ref}' not found. Shapes on this slide: {names}")


def _bgr(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    return r + g * 256 + b * 65536


def _size(pres):
    return pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight


def _show(slide=None):
    """Jump the PowerPoint window to the slide just changed so the user can watch."""
    try:
        if slide is not None:
            _app().ActiveWindow.View.GotoSlide(slide.SlideIndex)
    except Exception:
        pass
    time.sleep(STEP_DELAY)


def _parse_items(content):
    """Turn text/list into [(text, level)]. Leading tabs or 2-space groups mean sub-bullets."""
    lines = content.split("\n") if isinstance(content, str) else [str(x) for x in content]
    items = []
    for ln in lines:
        if not ln.strip():
            continue
        ws = ln[: len(ln) - len(ln.lstrip(" \t"))]
        level = min(5, 1 + ws.count("\t") + len(ws.replace("\t", "")) // 2)
        text = re.sub(r"^[-•*]\s+", "", ln.strip())
        items.append((text, level))
    return items


def _set_body(shape, content):
    items = _parse_items(content)
    tr = shape.TextFrame.TextRange
    tr.Text = "\r".join(t for t, _ in items)  # PowerPoint paragraphs are separated by \r
    for i, (_, level) in enumerate(items, start=1):
        if level > 1:
            tr.Paragraphs(i, 1).IndentLevel = level


def _placeholders(slide):
    out = []
    for i in range(1, slide.Shapes.Placeholders.Count + 1):
        sh = slide.Shapes.Placeholders.Item(i)
        out.append((sh, sh.PlaceholderFormat.Type))
    return out


def _format_shape(shape, bold=None, italic=None, underline=None, font_size=None, font_name=None,
                  font_color=None, align=None, fill_color=None):
    if shape.HasTextFrame:
        tr = shape.TextFrame.TextRange
        if bold is not None:
            tr.Font.Bold = -1 if bold else 0
        if italic is not None:
            tr.Font.Italic = -1 if italic else 0
        if underline is not None:
            tr.Font.Underline = -1 if underline else 0
        if font_size:
            tr.Font.Size = font_size
        if font_name:
            tr.Font.Name = font_name
        if font_color:
            tr.Font.Color.RGB = _bgr(font_color)
        if align:
            tr.ParagraphFormat.Alignment = _ALIGN[align]
    if fill_color:
        shape.Fill.Visible = -1
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = _bgr(fill_color)


def _set_notes(slide, notes):
    slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = notes


def _add_textbox_raw(slide, text, left, top, width, height):
    tb = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = horizontal
    tb.TextFrame.WordWrap = -1
    tb.TextFrame.TextRange.Text = text
    return tb


def _new_slide(pres, layout, index):
    idx = int(index) if index else pres.Slides.Count + 1
    key = str(layout or "title_and_content").lower().replace(" ", "_")
    if key in _LAYOUTS:
        return pres.Slides.Add(idx, _LAYOUTS[key])
    layouts = pres.SlideMaster.CustomLayouts
    for i in range(1, layouts.Count + 1):  # a layout name from the deck's own template
        cl = layouts.Item(i)
        if str(layout).lower() in cl.Name.lower():
            return pres.Slides.AddSlide(idx, cl)
    raise ValueError(f"Unknown layout '{layout}'. Use one of {list(_LAYOUTS)} or a name from list_layouts.")


def _build_slide(pres, layout=None, title=None, body=None, body_right=None, notes=None, index=None):
    W, H = _size(pres)
    slide = _new_slide(pres, layout, index)
    phs = _placeholders(slide)
    title_sh = next((sh for sh, t in phs if t in (1, 3)), None)
    bodies = [sh for sh, t in phs if t in (2, 4, 7)]
    if title:
        if title_sh is not None:
            title_sh.TextFrame.TextRange.Text = title
        else:
            tb = _add_textbox_raw(slide, title, 40, 20, W - 80, 60)
            _format_shape(tb, bold=True, font_size=32)
    for sh, content in zip(bodies, [body, body_right]):
        if content:
            _set_body(sh, content)
    for sh, _ in phs:  # remove placeholders left empty so no "Click to add text" boxes remain
        try:
            if sh.HasTextFrame and not sh.TextFrame.TextRange.Text.strip():
                sh.Delete()
        except Exception:
            pass
    if notes:
        try:
            _set_notes(slide, notes)
        except Exception:
            pass
    return slide


# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------
def open_presentation(path=None):
    app = _app()
    if path:
        path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        pres = app.Presentations.Open(path)
    else:
        pres = app.Presentations.Add()
    app.Activate()
    W, H = _size(pres)
    return {"presentation": pres.Name, "slides": pres.Slides.Count, "slide_size_points": [W, H]}


def get_presentation_info():
    pres = _pres()
    W, H = _size(pres)
    n = pres.Slides.Count
    lines = []
    for i in range(1, min(n, 60) + 1):
        s = pres.Slides.Item(i)
        title = s.Shapes.Title.TextFrame.TextRange.Text if s.Shapes.HasTitle else ""
        lines.append(f"{i} [{s.CustomLayout.Name}] {title[:80]} ({s.Shapes.Count} shapes)")
    return {
        "presentation": pres.Name,
        "path": pres.FullName,
        "slide_count": n,
        "slide_size_points": [W, H],
        "slides": lines,
    }


def get_slide(slide):
    """Everything on one slide: shapes (number, name, type, position, text) and speaker notes."""
    pres = _pres()
    s = _slide(pres, slide)
    shapes = []
    for i in range(1, s.Shapes.Count + 1):
        sh = s.Shapes.Item(i)
        info = {"shape": i, "name": sh.Name, "type": _SHAPE_TYPES.get(sh.Type, sh.Type),
                "left": round(sh.Left), "top": round(sh.Top),
                "width": round(sh.Width), "height": round(sh.Height)}
        if sh.HasTextFrame:
            text = sh.TextFrame.TextRange.Text.replace("\r", " | ")
            if text.strip():
                info["text"] = text[:300]
        if sh.HasTable:
            tbl = sh.Table
            rows = []
            for r in range(1, min(tbl.Rows.Count, 12) + 1):
                rows.append([tbl.Cell(r, c).Shape.TextFrame.TextRange.Text for c in range(1, min(tbl.Columns.Count, 8) + 1)])
            info["table"] = rows
        shapes.append(info)
    notes = ""
    try:
        notes = s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text
    except Exception:
        pass
    return {"slide": int(slide), "layout": s.CustomLayout.Name, "shapes": shapes, "notes": notes[:500]}


def list_layouts():
    pres = _pres()
    layouts = pres.SlideMaster.CustomLayouts
    return {
        "simple_layout_names": list(_LAYOUTS),
        "deck_layout_names": [layouts.Item(i).Name for i in range(1, layouts.Count + 1)],
    }


def add_slide(layout="title_and_content", title=None, body=None, body_right=None, notes=None, index=None):
    """Add one slide (at the end unless index is given)."""
    pres = _pres()
    s = _build_slide(pres, layout, title, body, body_right, notes, index)
    _show(s)
    return {"slide_added": s.SlideIndex, "total_slides": pres.Slides.Count}


def add_slides(slides):
    """Add many slides in one call. Each item: layout, title, body, body_right, notes."""
    pres = _pres()
    added = []
    for spec in slides:
        s = _build_slide(pres, spec.get("layout"), spec.get("title"), spec.get("body"),
                         spec.get("body_right"), spec.get("notes"))
        added.append(s.SlideIndex)
        _show(s)
    return {"slides_added": added, "total_slides": pres.Slides.Count}


def edit_slide_text(slide, title=None, body=None, body_right=None, notes=None):
    """Replace the title / body text / notes of an existing slide."""
    pres = _pres()
    W, H = _size(pres)
    s = _slide(pres, slide)
    phs = _placeholders(s)
    title_sh = next((sh for sh, t in phs if t in (1, 3)), None)
    bodies = [sh for sh, t in phs if t in (2, 4, 7)]
    if title:
        if title_sh is not None:
            title_sh.TextFrame.TextRange.Text = title
        else:
            _add_textbox_raw(s, title, 40, 20, W - 80, 60)
    for k, content in enumerate([body, body_right]):
        if not content:
            continue
        if k < len(bodies):
            _set_body(bodies[k], content)
        else:  # no body placeholder on this slide: use a text box instead
            tb = _add_textbox_raw(s, "", 40 + k * (W / 2), 110, (W - 100) / (2 if body_right else 1), H - 160)
            _set_body(tb, content)
    if notes is not None:
        _set_notes(s, notes)
    _show(s)
    return {"slide_edited": int(slide)}


def edit_shape_text(slide, shape, text):
    """Replace the text of one shape (number or name from get_slide)."""
    pres = _pres()
    s = _slide(pres, slide)
    sh = _shape(s, shape)
    if not sh.HasTextFrame:
        raise ValueError("That shape has no text. Use edit_table_cell for tables.")
    _set_body(sh, text)
    _show(s)
    return {"edited_shape": sh.Name}


def add_textbox(slide, text, left=40, top=120, width=None, height=60, font_size=None,
                bold=None, italic=None, font_color=None, align=None, fill_color=None):
    pres = _pres()
    W, H = _size(pres)
    s = _slide(pres, slide)
    tb = _add_textbox_raw(s, text, left, top, width or (W - 2 * left), height)
    _format_shape(tb, bold=bold, italic=italic, font_size=font_size, font_color=font_color,
                  align=align, fill_color=fill_color)
    _show(s)
    return {"textbox_added": tb.Name}


def add_picture(slide, path, left=None, top=None, width=None):
    """Add an image. Without left/top/width it is sized to fit and centered below the title."""
    path = os.path.abspath(os.path.expanduser(path))
    if not os.path.exists(path):
        raise FileNotFoundError(f"Image not found: {path}")
    pres = _pres()
    W, H = _size(pres)
    s = _slide(pres, slide)
    pic = s.Shapes.AddPicture(path, 0, -1, 0, 0, -1, -1)  # not linked, saved with file, native size
    pic.LockAspectRatio = -1
    w, h = pic.Width, pic.Height
    if width:
        pic.Width = float(width)
    else:
        scale = min(1.0, (W * 0.8) / w, (H * 0.65) / h)
        pic.Width = w * scale
    pic.Left = float(left) if left is not None else (W - pic.Width) / 2
    pic.Top = float(top) if top is not None else 110 + max(0, (H - 110 - pic.Height) / 2)
    _show(s)
    return {"picture_added": pic.Name}


def add_slide_table(slide, values, left=None, top=None, width=None, height=None,
                    header_fill=None, font_size=None):
    if not values or not isinstance(values[0], list):
        raise ValueError("values must be a list of rows, e.g. [['Item','Price'],['Tea',20]]")
    nrows = len(values)
    ncols = max(len(r) for r in values)
    if nrows * ncols > 200:
        raise ValueError("Table too large for a slide (max 200 cells). Split it across slides.")
    pres = _pres()
    W, H = _size(pres)
    s = _slide(pres, slide)
    shape = s.Shapes.AddTable(nrows, ncols, left if left is not None else 40,
                              top if top is not None else 120,
                              width or (W - 80), height or 30 * nrows)
    tbl = shape.Table
    fs = font_size or (18 if nrows <= 5 else 14 if nrows <= 8 else 12)
    if ncols > 5:
        fs = max(10, fs - 2)
    for r, row in enumerate(values, start=1):
        for c in range(1, ncols + 1):
            v = row[c - 1] if c - 1 < len(row) else ""
            tr = tbl.Cell(r, c).Shape.TextFrame.TextRange
            tr.Text = "" if v is None else str(v)
            tr.Font.Size = fs
    if header_fill:
        for c in range(1, ncols + 1):
            cell_shape = tbl.Cell(1, c).Shape
            cell_shape.Fill.Solid()
            cell_shape.Fill.ForeColor.RGB = _bgr(header_fill)
    _show(s)
    return {"table_added": shape.Name, "rows": nrows, "columns": ncols}


def edit_table_cell(slide, shape, row, column, text):
    pres = _pres()
    s = _slide(pres, slide)
    sh = _shape(s, shape)
    if not sh.HasTable:
        raise ValueError("That shape is not a table.")
    sh.Table.Cell(int(row), int(column)).Shape.TextFrame.TextRange.Text = str(text)
    _show(s)
    return {"cell_updated": [int(row), int(column)]}


def format_shape_text(slide, shape, bold=None, italic=None, underline=None, font_size=None,
                      font_name=None, font_color=None, align=None, fill_color=None):
    """Format a shape's text (and optionally its fill color)."""
    pres = _pres()
    s = _slide(pres, slide)
    sh = _shape(s, shape)
    _format_shape(sh, bold, italic, underline, font_size, font_name, font_color, align, fill_color)
    _show(s)
    return {"formatted_shape": sh.Name}


def set_slide_background(color, slide=None):
    """Set a solid background color for one slide, or for the whole deck (via the master) if slide is omitted."""
    pres = _pres()
    if slide:
        s = _slide(pres, slide)
        s.FollowMasterBackground = 0
        s.Background.Fill.Solid()
        s.Background.Fill.ForeColor.RGB = _bgr(color)
        _show(s)
        return {"background_set_for_slide": int(slide)}
    m = pres.SlideMaster
    m.Background.Fill.Solid()
    m.Background.Fill.ForeColor.RGB = _bgr(color)
    return {"background_set_for": "whole deck"}


def set_theme(heading_font=None, body_font=None, accent_color=None, background_color=None,
              title_color=None, text_color=None):
    """Deck-wide look via the slide master. Set text colors too when using a dark background."""
    pres = _pres()
    m = pres.SlideMaster
    applied, failed = [], []

    def attempt(label, fn):
        try:
            fn()
            applied.append(label)
        except Exception as e:  # noqa: BLE001
            failed.append(f"{label}: {e}")

    if heading_font:
        attempt("heading_font", lambda: setattr(m.Theme.ThemeFontScheme.MajorFont.Item(1), "Name", heading_font))
    if body_font:
        attempt("body_font", lambda: setattr(m.Theme.ThemeFontScheme.MinorFont.Item(1), "Name", body_font))
    if accent_color:
        attempt("accent_color", lambda: setattr(m.Theme.ThemeColorScheme.Colors(5), "RGB", _bgr(accent_color)))
    if background_color:
        def bg():
            m.Background.Fill.Solid()
            m.Background.Fill.ForeColor.RGB = _bgr(background_color)
        attempt("background_color", bg)
    if title_color:
        attempt("title_color", lambda: setattr(m.Shapes.Title.TextFrame.TextRange.Font.Color, "RGB", _bgr(title_color)))
    if text_color:
        attempt("text_color", lambda: setattr(m.Shapes.Placeholders.Item(2).TextFrame.TextRange.Font.Color, "RGB", _bgr(text_color)))
    _show(pres.Slides.Item(1) if pres.Slides.Count else None)
    return {"applied": applied, "failed": failed}


def move_slide(slide, to_position):
    pres = _pres()
    s = _slide(pres, slide)
    s.MoveTo(int(to_position))
    return {"moved": int(slide), "to": int(to_position)}


def duplicate_slide(slide):
    pres = _pres()
    s = _slide(pres, slide)
    s.Duplicate()
    return {"duplicated": int(slide), "total_slides": pres.Slides.Count}


def delete_slide(slide):
    pres = _pres()
    s = _slide(pres, slide)
    s.Delete()
    return {"deleted": int(slide), "total_slides": pres.Slides.Count}


def save_presentation(path=None, overwrite=False):
    """Save only when the user asks. Never overwrites an existing file unless overwrite=True."""
    app = _app()
    pres = _pres()
    if path:
        if not path.lower().endswith((".pptx", ".ppt")):
            path += ".pptx"
        if not os.path.isabs(path):
            path = os.path.join(os.path.expanduser("~"), "Documents", path)
        if os.path.exists(path) and not overwrite:
            raise FileExistsError(f"{path} already exists. Ask the user before overwriting.")
        app.DisplayAlerts = 1  # ppAlertsNone
        try:
            pres.SaveAs(path, 24)  # ppSaveAsOpenXMLPresentation (.pptx)
        finally:
            app.DisplayAlerts = 2  # ppAlertsAll
        return {"saved_to": path}
    if not pres.Path:
        raise ValueError("This presentation has never been saved. Provide a file name in 'path'.")
    pres.Save()
    return {"saved_to": pres.FullName}


def export_presentation_pdf(path, overwrite=False):
    pres = _pres()
    if not path.lower().endswith(".pdf"):
        path += ".pdf"
    if not os.path.isabs(path):
        path = os.path.join(os.path.expanduser("~"), "Documents", path)
    if os.path.exists(path) and not overwrite:
        raise FileExistsError(f"{path} already exists. Ask the user before overwriting.")
    try:
        pres.ExportAsFixedFormat(path, 2)  # ppFixedFormatTypePDF
    except Exception:
        pres.SaveCopyAs(path, 32)  # ppSaveAsPDF
    return {"pdf_saved_to": path}


TOOLS = {
    f.__name__: f
    for f in [
        open_presentation, get_presentation_info, get_slide, list_layouts, add_slide, add_slides,
        edit_slide_text, edit_shape_text, add_textbox, add_picture, add_slide_table, edit_table_cell,
        format_shape_text, set_slide_background, set_theme, move_slide, duplicate_slide,
        delete_slide, save_presentation, export_presentation_pdf,
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


_SLIDE = {"type": "integer", "description": "Slide number (1-based)."}
_SHAPE = {"type": "string", "description": "Shape number or name, as shown by get_slide."}
_COLOR = {"type": "string", "description": "Hex color like '#1F4E78'."}
_ALIGN_PROP = {"type": "string", "enum": ["left", "center", "right", "justify"]}
_LAYOUT = {"type": "string", "description": "title, title_and_content, section_header, two_content, title_only, blank, or a layout name from list_layouts."}
_BODY = {"type": "array", "items": {"type": "string"},
         "description": "Bullet lines. Start a line with two spaces or a tab for a sub-bullet. For layout 'title' this is the subtitle."}

TOOL_SCHEMAS = [
    _schema("open_presentation",
            "Open the real PowerPoint window. With 'path' opens that existing file; without it creates a new blank presentation.",
            {"path": {"type": "string", "description": "Full path of an existing .pptx (optional)."}}),
    _schema("get_presentation_info",
            "See the open presentation: slide count, slide size and a numbered list of slides with layout and title. Call before editing an existing deck.",
            {}),
    _schema("get_slide", "Inspect one slide: its shapes (number, name, type, position, text, table contents) and speaker notes.",
            {"slide": _SLIDE}, ["slide"]),
    _schema("list_layouts", "List the layouts available for add_slide.", {}),
    _schema("add_slide",
            "Add one slide with title, bullets and optional speaker notes. The first slide of a deck should use layout 'title'.",
            {"layout": _LAYOUT, "title": {"type": "string"}, "body": _BODY,
             "body_right": {"type": "array", "items": {"type": "string"}, "description": "Right column bullets for layout two_content."},
             "notes": {"type": "string", "description": "Speaker notes."},
             "index": {"type": "integer", "description": "Insert position; default is the end."}}),
    _schema("add_slides",
            "Add several slides in ONE call (preferred for building a deck). Each slide: layout, title, body (array of bullet strings), body_right, notes.",
            {"slides": {"type": "array", "items": {"type": "object", "properties": {
                "layout": _LAYOUT, "title": {"type": "string"}, "body": _BODY,
                "body_right": {"type": "array", "items": {"type": "string"}}, "notes": {"type": "string"}}}}},
            ["slides"]),
    _schema("edit_slide_text", "Replace the title, bullets and/or speaker notes of an existing slide.",
            {"slide": _SLIDE, "title": {"type": "string"}, "body": _BODY,
             "body_right": {"type": "array", "items": {"type": "string"}}, "notes": {"type": "string"}},
            ["slide"]),
    _schema("edit_shape_text", "Replace the text of one shape on a slide.",
            {"slide": _SLIDE, "shape": _SHAPE, "text": {"type": "string"}}, ["slide", "shape", "text"]),
    _schema("add_textbox", "Add a text box. Positions are in points (slide is typically 960 x 540).",
            {"slide": _SLIDE, "text": {"type": "string"}, "left": {"type": "number"}, "top": {"type": "number"},
             "width": {"type": "number"}, "height": {"type": "number"}, "font_size": {"type": "number"},
             "bold": {"type": "boolean"}, "italic": {"type": "boolean"}, "font_color": _COLOR,
             "align": _ALIGN_PROP, "fill_color": _COLOR},
            ["slide", "text"]),
    _schema("add_picture", "Add an image from a file path; by default it is sized to fit and centered.",
            {"slide": _SLIDE, "path": {"type": "string"}, "left": {"type": "number"},
             "top": {"type": "number"}, "width": {"type": "number"}}, ["slide", "path"]),
    _schema("add_slide_table", "Add a table to a slide. values is a list of rows; the first row is the header.",
            {"slide": _SLIDE,
             "values": {"type": "array", "items": {"type": "array", "items": {"type": ["string", "number", "boolean", "null"]}}},
             "left": {"type": "number"}, "top": {"type": "number"}, "width": {"type": "number"},
             "height": {"type": "number"}, "header_fill": _COLOR, "font_size": {"type": "number"}},
            ["slide", "values"]),
    _schema("edit_table_cell", "Change the text of one table cell (row and column are 1-based).",
            {"slide": _SLIDE, "shape": _SHAPE, "row": {"type": "integer"}, "column": {"type": "integer"},
             "text": {"type": "string"}},
            ["slide", "shape", "row", "column", "text"]),
    _schema("format_shape_text", "Format a shape's text and/or fill. Set only what should change.",
            {"slide": _SLIDE, "shape": _SHAPE, "bold": {"type": "boolean"}, "italic": {"type": "boolean"},
             "underline": {"type": "boolean"}, "font_size": {"type": "number"}, "font_name": {"type": "string"},
             "font_color": _COLOR, "align": _ALIGN_PROP, "fill_color": _COLOR},
            ["slide", "shape"]),
    _schema("set_slide_background", "Set a solid background color for one slide, or the whole deck when slide is omitted.",
            {"color": _COLOR, "slide": _SLIDE}, ["color"]),
    _schema("set_theme",
            "Set the deck-wide look through the slide master: fonts, accent color, background and title/body text colors. Call it right after opening a new presentation. With a dark background also set title_color and text_color to light colors.",
            {"heading_font": {"type": "string"}, "body_font": {"type": "string"}, "accent_color": _COLOR,
             "background_color": _COLOR, "title_color": _COLOR, "text_color": _COLOR}),
    _schema("move_slide", "Move a slide to a new position.", {"slide": _SLIDE, "to_position": {"type": "integer"}},
            ["slide", "to_position"]),
    _schema("duplicate_slide", "Duplicate a slide (the copy appears right after it).", {"slide": _SLIDE}, ["slide"]),
    _schema("delete_slide", "Delete a slide.", {"slide": _SLIDE}, ["slide"]),
    _schema("save_presentation",
            "Save the presentation. ONLY call when the user asks to save. Never overwrite an existing file unless the user explicitly agrees. A relative name is saved in the Documents folder.",
            {"path": {"type": "string"}, "overwrite": {"type": "boolean"}}),
    _schema("export_presentation_pdf", "Export the presentation as a PDF. Only when the user asks.",
            {"path": {"type": "string"}, "overwrite": {"type": "boolean"}}, ["path"]),
]


def call_tool(name: str, args: dict) -> str:
    return dispatch(_worker, TOOLS, name, args)
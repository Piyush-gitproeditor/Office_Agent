"""Live Excel tools. The agent drives the REAL Excel window through COM (xlwings).

Windows + desktop Excel required. All Excel calls run on ONE dedicated thread,
because COM objects can't be shared across the threads Streamlit creates.
"""

import json
import os
import queue
import threading
import time
from concurrent.futures import Future

import xlwings as xw

# Pause (seconds) after each visible step so you can follow along on screen.
STEP_DELAY = float(os.getenv("EXCEL_STEP_DELAY", "0.4"))
MAX_READ_ROWS = 200
MAX_RESULT_CHARS = 12000


# --------------------------------------------------------------------------
# Dedicated COM thread
# --------------------------------------------------------------------------
class _ExcelWorker:
    def __init__(self):
        self._q: queue.Queue = queue.Queue()
        threading.Thread(target=self._loop, daemon=True).start()

    def _loop(self):
        try:
            import pythoncom

            pythoncom.CoInitialize()
        except ImportError:
            pass
        while True:
            fn, fut = self._q.get()
            try:
                fut.set_result(fn())
            except Exception as e:  # noqa: BLE001 - report everything back to the caller
                fut.set_exception(e)

    def run(self, fn, timeout=180):
        fut: Future = Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)


_worker = _ExcelWorker()
_app = None  # only touched from the worker thread


# --------------------------------------------------------------------------
# Helpers (run on the worker thread)
# --------------------------------------------------------------------------
def _clean(addr: str) -> str:
    return addr.replace("$", "")


def _excel():
    """Return the running Excel app, starting a visible one if needed."""
    global _app
    if _app is not None:
        try:
            _app.books.count
        except Exception:  # user closed Excel
            _app = None
    if _app is None:
        _app = xw.App(visible=True, add_book=False)
    return _app


def _book():
    app = _excel()
    if app.books.count == 0:
        app.books.add()
    return app.books.active


def _sheet(name=None):
    wb = _book()
    if not name:
        return wb.sheets.active
    for s in wb.sheets:
        if s.name.lower() == str(name).lower():
            return s
    raise ValueError(f"Sheet '{name}' not found. Sheets: {[s.name for s in wb.sheets]}")


def _follow(rng):
    """Select the range so the user sees where the agent is working."""
    try:
        rng.sheet.activate()
        rng.select()
    except Exception:
        pass
    time.sleep(STEP_DELAY)


def _rgb(hex_color: str):
    h = hex_color.lstrip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def _to_2d(v):
    if not isinstance(v, list):
        return [[v]]
    if v and not isinstance(v[0], list):
        return [v]
    return v


# --------------------------------------------------------------------------
# Tools
# --------------------------------------------------------------------------
def open_workbook(path=None):
    app = _excel()
    if path:
        path = os.path.abspath(os.path.expanduser(path))
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        wb = app.books.open(path)
    else:
        wb = app.books.add()
    app.activate(steal_focus=True)
    return {"workbook": wb.name, "sheets": [s.name for s in wb.sheets]}


def get_workbook_info(preview_rows=15):
    app = _excel()
    wb = _book()
    ws = wb.sheets.active
    used = ws.used_range
    empty = used.count == 1 and used.value is None
    info = {
        "workbook": wb.name,
        "path": wb.fullname,
        "sheets": [s.name for s in wb.sheets],
        "active_sheet": ws.name,
        "used_range": None if empty else _clean(used.address),
    }
    try:
        info["selection"] = _clean(app.selection.address)
    except Exception:
        pass
    if not empty:
        last_row = used.last_cell.row
        last_col = min(used.last_cell.column, 30)
        n = min(int(preview_rows), last_row)
        info["preview_from_A1"] = ws.range((1, 1), (n, last_col)).options(ndim=2).value
    return info


def read_range(cell_range, sheet=None, include_formulas=False):
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    if rng.rows.count > MAX_READ_ROWS:
        rng = rng.resize(MAX_READ_ROWS)
    out = {"sheet": ws.name, "range": _clean(rng.address), "values": rng.options(ndim=2).value}
    if include_formulas:
        f = rng.formula
        out["formulas"] = [[f]] if isinstance(f, str) else [list(r) for r in f]
    return out


def write_cells(start_cell, values, sheet=None):
    ws = _sheet(sheet)
    values = _to_2d(values)
    nrows = len(values)
    ncols = max(len(r) for r in values)
    values = [list(r) + [None] * (ncols - len(r)) for r in values]
    top = ws.range(start_cell)
    target = top.resize(nrows, ncols)
    ws.activate()
    if 1 < nrows <= 30:  # typing effect: row by row so you can watch it build
        for i, row in enumerate(values):
            top.offset(i, 0).resize(1, ncols).value = [row]
            time.sleep(0.08)
    else:
        target.value = values
    target.columns.autofit()
    _follow(target)
    return {"sheet": ws.name, "written_range": _clean(target.address)}


def set_formula(cell_range, formula, sheet=None):
    """Set a formula. On a multi-cell range, references adjust like a fill-down."""
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    if not formula.startswith("="):
        formula = "=" + formula
    rng.formula = formula
    _follow(rng)
    first = ws.range(cell_range).options(ndim=2).value[0][0]
    return {"range": _clean(rng.address), "formula": formula, "first_cell_result": first}


def format_range(
    cell_range,
    sheet=None,
    bold=None,
    italic=None,
    font_size=None,
    font_color=None,
    fill_color=None,
    number_format=None,
    horizontal_align=None,
    wrap_text=None,
    borders=None,
    autofit_columns=None,
    column_width=None,
):
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    if bold is not None:
        rng.font.bold = bold
    if italic is not None:
        rng.font.italic = italic
    if font_size:
        rng.font.size = font_size
    if font_color:
        rng.font.color = _rgb(font_color)
    if fill_color:
        rng.color = _rgb(fill_color)
    if number_format:
        rng.number_format = number_format
    if horizontal_align:
        codes = {"left": -4131, "center": -4108, "right": -4152}
        rng.api.HorizontalAlignment = codes[horizontal_align]
    if wrap_text is not None:
        rng.api.WrapText = wrap_text
    if borders:
        rng.api.Borders.LineStyle = 1
    if column_width:
        rng.column_width = column_width
    if autofit_columns:
        rng.columns.autofit()
    _follow(rng)
    return {"formatted": _clean(rng.address)}


def add_sheet(name):
    wb = _book()
    ws = wb.sheets.add(name=name, after=wb.sheets[-1])
    ws.activate()
    return {"sheets": [s.name for s in wb.sheets]}


def rename_sheet(old_name, new_name):
    ws = _sheet(old_name)
    ws.name = new_name
    return {"sheets": [s.name for s in ws.book.sheets]}


_CHART_TYPES = {
    "column": "column_clustered",
    "bar": "bar_clustered",
    "line": "line",
    "line_markers": "line_markers",
    "pie": "pie",
    "scatter": "xy_scatter",
    "area": "area",
}


def create_chart(data_range, chart_type="column", title=None, sheet=None, position_cell=None):
    """data_range should include the header row/column so series get names."""
    ws = _sheet(sheet)
    src = ws.range(data_range)
    anchor = ws.range(position_cell) if position_cell else ws.range((1, src.last_cell.column + 2))
    ch = ws.charts.add(left=anchor.left, top=anchor.top, width=420, height=260)
    ch.set_source_data(src)
    ch.chart_type = _CHART_TYPES.get(chart_type, "column_clustered")
    if title:
        chart_obj = ch.api[1]
        chart_obj.HasTitle = True
        chart_obj.ChartTitle.Text = title
    ws.activate()
    time.sleep(STEP_DELAY)
    return {"chart_added": True, "type": chart_type, "source": _clean(src.address)}


def sort_range(cell_range, column_index=1, ascending=True, has_header=True, sheet=None):
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    key = rng.columns[int(column_index) - 1]
    rng.api.Sort(Key1=key.api, Order1=1 if ascending else 2, Header=1 if has_header else 2)
    _follow(rng)
    return {"sorted": _clean(rng.address), "by_column": column_index, "ascending": ascending}


def toggle_autofilter(cell_range, sheet=None):
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    rng.api.AutoFilter()
    _follow(rng)
    return {"autofilter_toggled_on": _clean(rng.address)}


def clear_range(cell_range, sheet=None, clear_formatting=False):
    ws = _sheet(sheet)
    rng = ws.range(cell_range)
    rng.clear() if clear_formatting else rng.clear_contents()
    _follow(rng)
    return {"cleared": _clean(rng.address)}


def save_workbook(path=None, overwrite=False):
    """Save only when the user asks. Never overwrites an existing file unless overwrite=True."""
    app = _excel()
    wb = _book()
    if path:
        if not path.lower().endswith((".xlsx", ".xlsm", ".xls")):
            path += ".xlsx"
        if not os.path.isabs(path):
            path = os.path.join(os.path.expanduser("~"), "Documents", path)
        if os.path.exists(path) and not overwrite:
            raise FileExistsError(f"{path} already exists. Ask the user before overwriting.")
        app.display_alerts = False
        try:
            wb.save(path)
        finally:
            app.display_alerts = True
        return {"saved_to": path}
    if not os.path.exists(wb.fullname):
        raise ValueError("This workbook has never been saved. Provide a file name in 'path'.")
    wb.save()
    return {"saved_to": wb.fullname}


TOOLS = {
    f.__name__: f
    for f in [
        open_workbook,
        get_workbook_info,
        read_range,
        write_cells,
        set_formula,
        format_range,
        add_sheet,
        rename_sheet,
        create_chart,
        sort_range,
        toggle_autofilter,
        clear_range,
        save_workbook,
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


_SHEET = {"type": "string", "description": "Sheet name. Defaults to the active sheet."}
_RANGE = {"type": "string", "description": "Excel address like 'A1:D10'."}
_CELL_VALUE = {"type": ["string", "number", "boolean", "null"]}

TOOL_SCHEMAS = [
    _schema(
        "open_workbook",
        "Open the real Excel window. With 'path' opens that existing file; without it creates a new blank workbook.",
        {"path": {"type": "string", "description": "Full path of an existing .xlsx file (optional)."}},
    ),
    _schema(
        "get_workbook_info",
        "See the current workbook: sheet names, used range, current selection and a preview of the top-left cells. Call this before editing an existing sheet.",
        {"preview_rows": {"type": "integer", "description": "Rows to preview (default 15)."}},
    ),
    _schema(
        "read_range",
        "Read values (and optionally formulas) from a range.",
        {"cell_range": _RANGE, "sheet": _SHEET, "include_formulas": {"type": "boolean"}},
        ["cell_range"],
    ),
    _schema(
        "write_cells",
        "Write a 2D grid of values starting at start_cell. Text starting with '=' becomes a formula. Example values: [['Name','Marks'],['Asha',91]].",
        {
            "start_cell": {"type": "string", "description": "Top-left cell, e.g. 'A1'."},
            "values": {"type": "array", "items": {"type": "array", "items": _CELL_VALUE}},
            "sheet": _SHEET,
        },
        ["start_cell", "values"],
    ),
    _schema(
        "set_formula",
        "Set one formula on a cell or range. On a multi-cell range, relative references adjust per row like fill-down, e.g. cell_range='D2:D20', formula='=B2*C2'.",
        {"cell_range": _RANGE, "formula": {"type": "string"}, "sheet": _SHEET},
        ["cell_range", "formula"],
    ),
    _schema(
        "format_range",
        "Format a range. Colors are hex like '#1F4E78'. number_format examples: '0.00', '#,##0', '0%', 'dd-mmm-yyyy', '₹#,##0'.",
        {
            "cell_range": _RANGE,
            "sheet": _SHEET,
            "bold": {"type": "boolean"},
            "italic": {"type": "boolean"},
            "font_size": {"type": "number"},
            "font_color": {"type": "string"},
            "fill_color": {"type": "string"},
            "number_format": {"type": "string"},
            "horizontal_align": {"type": "string", "enum": ["left", "center", "right"]},
            "wrap_text": {"type": "boolean"},
            "borders": {"type": "boolean", "description": "Thin borders on all cells."},
            "autofit_columns": {"type": "boolean"},
            "column_width": {"type": "number"},
        },
        ["cell_range"],
    ),
    _schema("add_sheet", "Add a new sheet at the end.", {"name": {"type": "string"}}, ["name"]),
    _schema(
        "rename_sheet",
        "Rename a sheet.",
        {"old_name": {"type": "string"}, "new_name": {"type": "string"}},
        ["old_name", "new_name"],
    ),
    _schema(
        "create_chart",
        "Create a chart from a range. Include the header row so series are named.",
        {
            "data_range": _RANGE,
            "chart_type": {"type": "string", "enum": list(_CHART_TYPES)},
            "title": {"type": "string"},
            "sheet": _SHEET,
            "position_cell": {"type": "string", "description": "Where to place the chart's top-left corner."},
        },
        ["data_range"],
    ),
    _schema(
        "sort_range",
        "Sort a range by one of its columns (1 = first column of the range).",
        {
            "cell_range": _RANGE,
            "column_index": {"type": "integer"},
            "ascending": {"type": "boolean"},
            "has_header": {"type": "boolean"},
            "sheet": _SHEET,
        },
        ["cell_range"],
    ),
    _schema(
        "toggle_autofilter",
        "Turn filter dropdown arrows on or off for a range (include the header row).",
        {"cell_range": _RANGE, "sheet": _SHEET},
        ["cell_range"],
    ),
    _schema(
        "clear_range",
        "Clear cell contents (and formatting if clear_formatting is true).",
        {"cell_range": _RANGE, "sheet": _SHEET, "clear_formatting": {"type": "boolean"}},
        ["cell_range"],
    ),
    _schema(
        "save_workbook",
        "Save the workbook. ONLY call when the user asks to save. Never overwrite an existing file unless the user explicitly agrees. A relative name is saved in the Documents folder.",
        {"path": {"type": "string"}, "overwrite": {"type": "boolean"}},
    ),
]


# --------------------------------------------------------------------------
# Entry point used by the agent loop
# --------------------------------------------------------------------------
def call_tool(name: str, args: dict) -> str:
    fn = TOOLS.get(name)
    if fn is None:
        return json.dumps({"error": f"Unknown tool '{name}'. Available: {list(TOOLS)}"})
    try:
        result = _worker.run(lambda: fn(**args))
    except TypeError as e:
        return json.dumps({"error": f"Bad arguments for {name}: {e}"})
    except Exception as e:  # noqa: BLE001
        return json.dumps({"error": f"{type(e).__name__}: {e}"})
    text = json.dumps(result, default=str, ensure_ascii=False)
    return text[:MAX_RESULT_CHARS]
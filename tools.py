"""One registry for all Office tools (Excel + Word + PowerPoint)."""

import json

from tools_excel import TOOL_SCHEMAS as _EXCEL_SCHEMAS
from tools_excel import TOOLS as _EXCEL_TOOLS
from tools_excel import call_tool as _excel_call
from Tools_ppt import TOOL_SCHEMAS as _PPT_SCHEMAS
from Tools_ppt import TOOLS as _PPT_TOOLS
from Tools_ppt import call_tool as _ppt_call
from tools_word import TOOL_SCHEMAS as _WORD_SCHEMAS
from tools_word import TOOLS as _WORD_TOOLS
from tools_word import call_tool as _word_call

TOOL_SCHEMAS = _EXCEL_SCHEMAS + _WORD_SCHEMAS + _PPT_SCHEMAS

_ROUTES = [(_EXCEL_TOOLS, _excel_call), (_WORD_TOOLS, _word_call), (_PPT_TOOLS, _ppt_call)]


def call_tool(name: str, args: dict) -> str:
    for tools, call in _ROUTES:
        if name in tools:
            return call(name, args)
    available = [n for tools, _ in _ROUTES for n in tools]
    return json.dumps({"error": f"Unknown tool '{name}'. Available: {available}"})
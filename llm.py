"""LLM layer: Nemotron on Nebius Token Factory (OpenAI-compatible API) + tool-calling loop."""

import json
import os
from typing import Iterator

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()
# Model Name = nemotron-3-super-120b-a12b  
BASE_URL = os.getenv("NEBIUS_BASE_URL", "https://api.tokenfactory.nebius.com/v1/")
DEFAULT_MODEL = os.getenv("NEBIUS_MODEL", "")

MAX_HISTORY_MESSAGES = 40  # keeps long chats inside the context window
MAX_STEPS = 30  # max tool-calling rounds per user message

SYSTEM_PROMPT = """You are an assistant that works inside the user's REAL Microsoft Excel, Word and PowerPoint windows on Windows.
The user watches the app on screen while you work, so build things step by step.

Tools available now: live Excel, Word and PowerPoint tools. Windows diagnostics
is not connected yet; say so plainly if asked for them.

Rules:
- For a new spreadsheet request, call open_workbook (blank) unless the user names a file.
- For an existing sheet, call get_workbook_info first to see the layout before editing.
- If the user's message says a spreadsheet is attached with a saved path, call open_workbook with that exact path,
  then get_workbook_info before editing. Edits happen on that uploaded copy, not the user's original.
- For a new Word document, call open_document (blank). If the user's message says a Word document is attached with a saved path,
  call open_document with that exact path, then get_document_info before editing; edits happen on the uploaded copy.
- In Word, build content in order and use styles (Title, Heading 1/2, Normal, List Bullet) instead of manual big fonts.
  To change existing text, use edit_paragraph / replace_text / format_text with paragraph numbers from get_document_info.
- For a new PowerPoint, call open_presentation (blank), then set_theme for fonts/colors, then build the deck with add_slides
  (one call for many slides; the first slide uses layout 'title'). If the user's message says a presentation is attached with a
  saved path, call open_presentation with that exact path, then get_presentation_info and get_slide before editing; edits happen on the uploaded copy.
- Slides: one idea per slide, a short title, 3-6 short bullets (under about 12 words each); put extra detail in speaker notes.
  Never cram text: split into more slides instead. Slide numbers are 1-based; re-check with get_presentation_info after adding, moving or deleting slides.
- Work in logical steps: headers, data, formulas, formatting, then charts.
- Prefer Excel formulas over typing calculated numbers by hand.
- NEVER save unless the user asks. When saving, never overwrite an existing file unless the user explicitly says so.
- If a tool returns an error, fix the call and retry once; if it still fails, explain the problem.
- Never claim you did something you did not do with a tool.
- When finished, give a short summary of what is now in the workbook."""


def get_client() -> OpenAI:
    api_key = os.getenv("NEBIUS_API_KEY")
    if not api_key:
        raise RuntimeError("NEBIUS_API_KEY is missing. Add it to your .env file.")
    return OpenAI(base_url=BASE_URL, api_key=api_key)


def _trim(history: list[dict]) -> list[dict]:
    """Keep recent messages, always starting at a user message so tool calls never get orphaned."""
    msgs = history[-MAX_HISTORY_MESSAGES:]
    while msgs and msgs[0]["role"] != "user":
        msgs = msgs[1:]
    return msgs


def run_agent(
    history: list[dict],
    model: str | None = None,
    temperature: float = 0.3,
    max_tokens: int = 2048,
) -> Iterator[dict]:
    """Run one user turn. Appends assistant/tool messages to `history` in place.

    Yields events:
      {"type": "text", "text": str}
      {"type": "tool_call", "name": str, "args": dict}
      {"type": "tool_result", "name": str, "result": str}
    """
    from tools import TOOL_SCHEMAS, call_tool  # imported lazily (needs Windows + Office)

    model = model or DEFAULT_MODEL
    if not model:
        raise RuntimeError("No model set. Put the Nemotron model ID in NEBIUS_MODEL in .env.")
    client = get_client()

    for step in range(MAX_STEPS):
        messages = [{"role": "system", "content": SYSTEM_PROMPT}] + _trim(history)
        stream = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=TOOL_SCHEMAS,
            tool_choice="auto",
            temperature=temperature,
            max_tokens=max_tokens,
            stream=True,
        )

        content = ""
        calls: dict[int, dict] = {}
        for chunk in stream:
            if not chunk.choices:
                continue
            delta = chunk.choices[0].delta
            if delta.content:
                content += delta.content
                yield {"type": "text", "text": delta.content}
            for tc in delta.tool_calls or []:
                slot = calls.setdefault(tc.index, {"id": "", "name": "", "args": ""})
                if tc.id:
                    slot["id"] = tc.id
                if tc.function:
                    if tc.function.name:
                        slot["name"] = tc.function.name
                    if tc.function.arguments:
                        slot["args"] += tc.function.arguments

        if not calls:  # plain answer: the turn is finished
            history.append({"role": "assistant", "content": content})
            return

        ordered = [calls[i] for i in sorted(calls)]
        for i, c in enumerate(ordered):
            c["id"] = c["id"] or f"call_{step}_{i}"
        history.append(
            {
                "role": "assistant",
                "content": content,
                "tool_calls": [
                    {
                        "id": c["id"],
                        "type": "function",
                        "function": {"name": c["name"], "arguments": c["args"] or "{}"},
                    }
                    for c in ordered
                ],
            }
        )

        for c in ordered:
            try:
                args = json.loads(c["args"] or "{}")
            except json.JSONDecodeError:
                args = None
            yield {"type": "tool_call", "name": c["name"], "args": args if args is not None else c["args"]}
            if args is None:
                result = json.dumps({"error": "Arguments were not valid JSON. Retry the call."})
            else:
                result = call_tool(c["name"], args)
            history.append({"role": "tool", "tool_call_id": c["id"], "content": result})
            yield {"type": "tool_result", "name": c["name"], "result": result}

    note = "I stopped because I hit the step limit for one request. Tell me to continue if needed."
    history.append({"role": "assistant", "content": note})
    yield {"type": "text", "text": note}


if __name__ == "__main__":
    # Terminal test:  python llm.py
    hist: list[dict] = []
    print("Terminal chat test. Ctrl+C to quit.\n")
    while True:
        text = input("You: ").strip()
        if not text:
            continue
        hist.append({"role": "user", "content": text})
        print("Agent: ", end="", flush=True)
        for ev in run_agent(hist):
            if ev["type"] == "text":
                print(ev["text"], end="", flush=True)
            elif ev["type"] == "tool_call":
                print(f"\n  [tool] {ev['name']} {json.dumps(ev['args'], ensure_ascii=False)}")
            elif ev["type"] == "tool_result":
                print(f"  [result] {ev['result'][:200]}")
        print("\n")
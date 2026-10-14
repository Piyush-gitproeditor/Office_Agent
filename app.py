"""Chat UI for the Microsoft tasks agent.  Run with:  streamlit run app.py"""

import json
from pathlib import Path

import streamlit as st

from llm import DEFAULT_MODEL, run_agent


# --- File uploads -----------------------------------------------------------
UPLOAD_DIR = Path(__file__).parent / "uploads"
SHEET_TYPES = {".xlsx", ".xlsm", ".xls", ".csv"}  # opened live in Excel
DOC_TYPES = {".docx", ".doc"}  # opened live in Word
PPT_TYPES = {".pptx", ".ppt"}  # opened live in PowerPoint
TEXT_TYPES = {".txt", ".md", ".json"}  # contents are sent to the model
MAX_TEXT_CHARS = 20000
ACCEPTED = [t.lstrip(".") for t in sorted(SHEET_TYPES | DOC_TYPES | PPT_TYPES | TEXT_TYPES)]


def save_upload(f) -> str:
    """Save an uploaded file into ./uploads (never overwriting) and return its full path."""
    UPLOAD_DIR.mkdir(exist_ok=True)
    dest = UPLOAD_DIR / Path(f.name).name
    stem, suffix, n = dest.stem, dest.suffix, 1
    while dest.exists():
        dest = UPLOAD_DIR / f"{stem}_{n}{suffix}"
        n += 1
    dest.write_bytes(f.getvalue())
    return str(dest.resolve())


def build_message(submission) -> str:
    """Turn the chat box content (text + attached files) into one message for the model."""
    text = (submission["text"] or "").strip()
    parts = [text or "Please look at the attached file(s)."]
    for f in submission["files"]:
        ext = Path(f.name).suffix.lower()
        if ext in SHEET_TYPES:
            path = save_upload(f)
            parts.append(f"[Attached spreadsheet '{f.name}'. Uploaded copy saved at: {path}]")
        elif ext in DOC_TYPES:
            path = save_upload(f)
            parts.append(f"[Attached Word document '{f.name}'. Uploaded copy saved at: {path}]")
        elif ext in PPT_TYPES:
            path = save_upload(f)
            parts.append(f"[Attached PowerPoint presentation '{f.name}'. Uploaded copy saved at: {path}]")
        else:
            body = f.getvalue().decode("utf-8", errors="replace")
            cut = " (truncated)" if len(body) > MAX_TEXT_CHARS else ""
            parts.append(f"[Attached text file '{f.name}'{cut}]\n{body[:MAX_TEXT_CHARS]}\n[end of file]")
    return "\n\n".join(parts)

st.set_page_config(page_title="Microsoft Tasks Agent", page_icon="🗂️")
st.title("🗂️ Microsoft Tasks Agent")
st.caption("Live Excel, Word and PowerPoint · Windows diagnostics coming next")

if "history" not in st.session_state:
    st.session_state.history = []  # full API-format history, including tool messages

with st.sidebar:
    st.header("Settings")
    model = st.text_input("Model ID", value=DEFAULT_MODEL)
    temperature = st.slider("Temperature", 0.0, 1.0, 0.3, 0.05)
    if st.button("Clear chat"):
        st.session_state.history = []
        st.rerun()


def show_tool(name, args, result=None):
    with st.expander(f"🔧 {name}", expanded=False):
        st.code(json.dumps(args, indent=2, ensure_ascii=False) if not isinstance(args, str) else args, language="json")
        if result is not None:
            st.code(result[:1500], language="json")


def render_assistant_block(block):
    results = {m["tool_call_id"]: m["content"] for m in block if m["role"] == "tool"}
    for m in block:
        if m["role"] != "assistant":
            continue
        if m.get("content"):
            st.markdown(m["content"])
        for tc in m.get("tool_calls", []):
            fn = tc["function"]
            try:
                args = json.loads(fn["arguments"] or "{}")
            except json.JSONDecodeError:
                args = fn["arguments"]
            show_tool(fn["name"], args, results.get(tc["id"]))


# Re-draw the conversation: user messages, then each assistant turn as one bubble
h = st.session_state.history
i = 0
while i < len(h):
    if h[i]["role"] == "user":
        with st.chat_message("user"):
            st.markdown(h[i]["content"])
        i += 1
    else:
        j = i
        while j < len(h) and h[j]["role"] != "user":
            j += 1
        with st.chat_message("assistant"):
            render_assistant_block(h[i:j])
        i = j

submission = st.chat_input(
    "Describe what to do in Excel, Word or PowerPoint, or attach a file...",
    accept_file="multiple",
    file_type=ACCEPTED,
)
if submission:
    prompt = build_message(submission)
    n_before = len(st.session_state.history)
    st.session_state.history.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        text_slot, buffer, current_tool = None, "", None
        try:
            for ev in run_agent(st.session_state.history, model=model, temperature=temperature):
                if ev["type"] == "text":
                    if text_slot is None:
                        text_slot, buffer = st.empty(), ""
                    buffer += ev["text"]
                    text_slot.markdown(buffer)
                elif ev["type"] == "tool_call":
                    text_slot = None
                    current_tool = st.status(f"🔧 {ev['name']}", expanded=False)
                    with current_tool:
                        a = ev["args"]
                        st.code(json.dumps(a, indent=2, ensure_ascii=False) if not isinstance(a, str) else a, language="json")
                elif ev["type"] == "tool_result" and current_tool is not None:
                    with current_tool:
                        st.code(ev["result"][:1500], language="json")
                    current_tool.update(state="complete")
        except Exception as e:
            st.error(f"Something went wrong: {e}")
            del st.session_state.history[n_before:]  # drop this unfinished turn
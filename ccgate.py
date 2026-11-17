"""
===============================================================================
ccgate.py — Claude Code Gateway
===============================================================================
A self-hosted protocol-translating reverse proxy that lets Claude Code talk to
ANY OpenAI-compatible backend provider (OpenRouter, Groq, NVIDIA AI Studio,
Hugging Face, Ollama, etc.).

Architecture:
    Claude Code ──► localhost:5005 ──► ccgate ──► OpenRouter / Groq / NVIDIA
                   (Anthropic API)    translates    (OpenAI API)

Usage:
    python ccgate.py                                # Interactive CLI
    python ccgate.py --provider groq --model openai/gpt-oss-20b
    python ccgate.py --port 8082                    # Custom port
    python ccgate.py setup                          # Write Claude Code config
    python ccgate.py doctor                         # Test all providers

Features:
    • Translates Anthropic Messages API ↔ OpenAI Chat Completions API
    • Full SSE streaming with strictly compliant real-time event translation
    • Tool call / function call bidirectional translation (stream & sync)
    • Integrates with providers.json and keys.xlsx for quick selection
    • One-command Claude Code setup
    • CORS support for any client
    • Zero external dependencies (pure Python 3.10+ standard library)
===============================================================================
"""

import sys
import os
import json
import time
import uuid
import ssl
import argparse
import signal
import threading
import traceback
import urllib.request
import urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from socketserver import ThreadingMixIn
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, Tuple, List

# Reconfigure stdout/stderr on Windows to avoid UnicodeEncodeError on cp1252/cp437
if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# =============================================================================
# CONSTANTS & PATHS
# =============================================================================

VERSION = "1.1.0"
DEFAULT_PORT = 5005
DEFAULT_HOST = "127.0.0.1"
PROJECT_DIR = Path(__file__).resolve().parent
PROVIDERS_FILE = PROJECT_DIR / "providers.json"
CLAUDE_SETTINGS = Path.home() / ".claude" / "settings.json"


# =============================================================================
# TERMINAL COLORS
# =============================================================================

class C:
    """ANSI escape codes for colored terminal output."""
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    DIM    = "\033[2m"
    RED    = "\033[91m"
    GREEN  = "\033[92m"
    YELLOW = "\033[93m"
    BLUE   = "\033[94m"
    MAGENTA= "\033[95m"
    CYAN   = "\033[96m"
    WHITE  = "\033[97m"
    BG_RED = "\033[41m"
    BG_GRN = "\033[42m"
    BG_BLU = "\033[44m"

    @staticmethod
    def disable():
        for attr in dir(C):
            if attr.isupper() and not attr.startswith("_"):
                setattr(C, attr, "")


# Disable colors if not a real terminal
if not sys.stdout.isatty():
    C.disable()


# =============================================================================
# DATA LOADING (providers.json & keys.xlsx)
# =============================================================================

def load_providers_json() -> Dict[str, Any]:
    """Load provider presets from providers.json."""
    if not PROVIDERS_FILE.exists():
        return {}
    try:
        data = json.loads(PROVIDERS_FILE.read_text(encoding="utf-8"))
        return data.get("providers", {})
    except Exception:
        return {}


def find_key_file() -> Path:
    """Find keys.xlsx in known locations."""
    for p in [
        Path(r"C:\My Space\Data\keys.xlsx"),
        PROJECT_DIR / "keys.xlsx",
        Path.cwd() / "keys.xlsx",
    ]:
        if p.exists():
            return p
    return Path(r"C:\My Space\Data\keys.xlsx")


def load_keys_from_excel() -> Dict[str, Dict[str, str]]:
    """
    Load API keys from keys.xlsx. Dual-method: openpyxl, then stdlib XML parser.
    Returns {nickname: {provider, key}}.
    """
    kf = find_key_file()
    if not kf.exists():
        return {}

    # Method 1: openpyxl
    try:
        import openpyxl
        wb = openpyxl.load_workbook(kf, read_only=True, data_only=True)
        ws = wb.active
        hdrs = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]
        pc = hdrs.index("provider") + 1 if "provider" in hdrs else 1
        nc = hdrs.index("nickname") + 1 if "nickname" in hdrs else 2
        kc = hdrs.index("api key") + 1 if "api key" in hdrs else (hdrs.index("key") + 1 if "key" in hdrs else 3)
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < nc:
                continue
            nick = str(row[nc - 1] or "").strip()
            if not nick:
                continue
            prov = str(row[pc - 1] or "").strip() if len(row) >= pc else ""
            key = str(row[kc - 1] or "").strip() if len(row) >= kc else ""
            if key:
                if key.lower().startswith("bearer "):
                    key = key[7:].strip()
                result[nick] = {"provider": prov, "key": key}
        wb.close()
        if result:
            return result
    except Exception:
        pass

    # Method 2: Pure stdlib (zipfile + xml.etree)
    try:
        import zipfile
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(kf, "r") as z:
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in root.findall(f".//{ns}si"):
                    t = si.find(f".//{ns}t")
                    shared.append(t.text if t is not None and t.text else "")
            sheets = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
            if not sheets:
                return {}
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            root = ET.fromstring(z.read(sheets[0]))
            rows = []
            for row_el in root.findall(f".//{ns}row"):
                vals = {}
                for cell in row_el.findall(f"{ns}c"):
                    ref = cell.attrib.get("r", "")
                    col = "".join(filter(str.isalpha, ref))
                    t = cell.attrib.get("t", "")
                    v = cell.find(f"{ns}v")
                    val = v.text if v is not None else ""
                    if t == "s" and val.isdigit():
                        idx = int(val)
                        val = shared[idx] if idx < len(shared) else val
                    vals[col] = val
                rows.append(vals)
            if not rows:
                return {}
            hdr = rows[0]
            cm = {}
            for col, txt in hdr.items():
                low = str(txt).strip().lower()
                if "provider" in low:
                    cm["provider"] = col
                elif "nickname" in low:
                    cm["nickname"] = col
                elif "key" in low:
                    cm["key"] = col
            pc, nc, kc = cm.get("provider", "A"), cm.get("nickname", "B"), cm.get("key", "C")
            result = {}
            for r in rows[1:]:
                nick = str(r.get(nc, "")).strip()
                prov = str(r.get(pc, "")).strip()
                key = str(r.get(kc, "")).strip()
                if nick and key:
                    if key.lower().startswith("bearer "):
                        key = key[7:].strip()
                    result[nick] = {"provider": prov, "key": key}
            return result
    except Exception:
        return {}


def find_key_for_provider(provider_name: str, saved_keys: Dict[str, Dict[str, str]]) -> str:
    """Find the best matching API key for a provider name."""
    p_low = provider_name.strip().lower()

    # Exact match on provider column
    for nick, data in saved_keys.items():
        if data.get("provider", "").strip().lower() == p_low:
            return data.get("key", "")

    # Substring match on provider column
    for nick, data in saved_keys.items():
        prov = data.get("provider", "").strip().lower()
        if p_low in prov or prov in p_low:
            return data.get("key", "")

    # Match on nickname
    for nick, data in saved_keys.items():
        n_low = nick.strip().lower()
        if p_low in n_low or n_low in p_low:
            return data.get("key", "")

    return ""


# =============================================================================
# URL BUILDER
# =============================================================================

def build_chat_url(base_url: str) -> str:
    """Build the /chat/completions endpoint from a base URL."""
    u = base_url.strip().rstrip("/")
    if u.endswith("/chat/completions"):
        return u
    if u.endswith("/v1"):
        return f"{u}/chat/completions"
    if u.endswith("/api"):
        return f"{u}/v1/chat/completions"
    return f"{u}/v1/chat/completions"


# =============================================================================
# PROTOCOL TRANSLATION: Anthropic Request → OpenAI Request
# =============================================================================

def translate_messages_to_openai(messages: List[Dict]) -> List[Dict]:
    """Convert Anthropic message format to OpenAI message format."""
    result = []
    for msg in messages:
        role = msg.get("role", "user")
        content = msg.get("content")

        # --- assistant message with content blocks ---
        if role == "assistant" and isinstance(content, list):
            text_parts = []
            tool_calls = []
            for block in content:
                if not isinstance(block, dict):
                    continue
                btype = block.get("type", "")
                if btype == "text":
                    text_parts.append(block.get("text", ""))
                elif btype == "tool_use":
                    inp = block.get("input")
                    if inp is None:
                        args_str = "{}"
                    elif isinstance(inp, dict):
                        args_str = json.dumps(inp)
                    elif isinstance(inp, str):
                        args_str = inp
                    else:
                        args_str = json.dumps(inp)

                    tool_calls.append({
                        "id": block.get("id", f"call_{uuid.uuid4().hex[:24]}"),
                        "type": "function",
                        "function": {
                            "name": block.get("name", ""),
                            "arguments": args_str,
                        },
                    })
                # Skip 'thinking' blocks — stripped for OpenAI backends

            openai_msg: Dict[str, Any] = {"role": "assistant"}
            if text_parts:
                openai_msg["content"] = "\n".join(text_parts)
            elif not tool_calls:
                # If only thinking blocks were present, OpenAI requires non-null content
                openai_msg["content"] = ""
            else:
                openai_msg["content"] = None

            if tool_calls:
                openai_msg["tool_calls"] = tool_calls
            result.append(openai_msg)
            continue

        # --- user message with content blocks ---
        if role == "user" and isinstance(content, list):
            tool_results = []
            other_parts = []
            for block in content:
                if not isinstance(block, dict):
                    continue
                btype = block.get("type", "")
                if btype == "tool_result":
                    tc = block.get("content", "")
                    if isinstance(tc, list):
                        tc = "\n".join(
                            p.get("text", "") for p in tc if isinstance(p, dict) and p.get("type") == "text"
                        )
                    elif not isinstance(tc, str):
                        tc = json.dumps(tc)
                    tool_results.append({
                        "role": "tool",
                        "tool_call_id": block.get("tool_use_id", ""),
                        "content": str(tc),
                    })
                elif btype == "text":
                    other_parts.append({"type": "text", "text": block.get("text", "")})
                elif btype == "image":
                    src = block.get("source", {})
                    if src.get("type") == "base64":
                        url = f"data:{src.get('media_type', 'image/png')};base64,{src.get('data', '')}"
                        other_parts.append({"type": "image_url", "image_url": {"url": url}})

            # Emit tool results first (OpenAI requires role 'tool' immediately following assistant tool_calls)
            for tr in tool_results:
                result.append(tr)

            # Emit remaining content as user message
            if other_parts:
                if len(other_parts) == 1 and other_parts[0]["type"] == "text":
                    result.append({"role": "user", "content": other_parts[0]["text"]})
                else:
                    result.append({"role": "user", "content": other_parts})
            elif not tool_results:
                result.append({"role": "user", "content": ""})
            continue

        # --- simple string content ---
        if isinstance(content, str):
            result.append({"role": role, "content": content})
        elif content is None:
            result.append({"role": role, "content": ""})
        else:
            result.append({"role": role, "content": str(content)})

    return result


def translate_tools_to_openai(tools: List[Dict]) -> List[Dict]:
    """Convert Anthropic tool definitions to OpenAI format."""
    result = []
    for tool in tools:
        if not isinstance(tool, dict):
            continue
        result.append({
            "type": "function",
            "function": {
                "name": tool.get("name", ""),
                "description": tool.get("description", ""),
                "parameters": tool.get("input_schema", {}),
            },
        })
    return result


def translate_tool_choice_to_openai(tc: Any) -> Any:
    """Convert Anthropic tool_choice to OpenAI format."""
    if not isinstance(tc, dict):
        return tc
    t = tc.get("type", "auto")
    if t == "auto":
        return "auto"
    if t == "any":
        return "required"
    if t == "tool":
        return {"type": "function", "function": {"name": tc.get("name", "")}}
    return "auto"


def build_openai_request(anthropic_body: Dict, backend_model: str) -> Dict:
    """Translate a full Anthropic Messages API request to OpenAI Chat Completions."""
    messages = []

    # 1. System prompt → system message
    system = anthropic_body.get("system")
    if system:
        if isinstance(system, str):
            messages.append({"role": "system", "content": system})
        elif isinstance(system, list):
            text = "\n".join(
                b.get("text", "") for b in system if isinstance(b, dict) and b.get("type") == "text"
            )
            if text:
                messages.append({"role": "system", "content": text})

    # 2. Translate messages
    messages.extend(translate_messages_to_openai(anthropic_body.get("messages", [])))

    # 3. Build OpenAI payload
    openai_req: Dict[str, Any] = {
        "model": backend_model,
        "messages": messages,
    }

    # Transfer compatible parameters
    if "max_tokens" in anthropic_body:
        openai_req["max_tokens"] = anthropic_body["max_tokens"]
    if "temperature" in anthropic_body:
        openai_req["temperature"] = anthropic_body["temperature"]
    if "top_p" in anthropic_body:
        openai_req["top_p"] = anthropic_body["top_p"]
    if "stop_sequences" in anthropic_body:
        openai_req["stop"] = anthropic_body["stop_sequences"]
    if anthropic_body.get("stream"):
        openai_req["stream"] = True
        openai_req["stream_options"] = {"include_usage": True}

    # Tools
    tools = anthropic_body.get("tools")
    if tools:
        openai_req["tools"] = translate_tools_to_openai(tools)
    tc = anthropic_body.get("tool_choice")
    if tc:
        openai_req["tool_choice"] = translate_tool_choice_to_openai(tc)

    return openai_req


# =============================================================================
# PROTOCOL TRANSLATION: OpenAI Response → Anthropic Response (Non-Streaming)
# =============================================================================

FINISH_REASON_MAP = {
    "stop": "end_turn",
    "length": "max_tokens",
    "tool_calls": "tool_use",
    "content_filter": "end_turn",
    "function_call": "tool_use",
}


def translate_openai_response(openai_resp: Dict, model_alias: str) -> Dict:
    """Convert an OpenAI chat.completion response to Anthropic Messages format."""
    choice = {}
    if openai_resp.get("choices"):
        choice = openai_resp["choices"][0]

    message = choice.get("message", {})
    finish = choice.get("finish_reason", "stop")

    # Build content blocks
    content_blocks = []

    # Text / reasoning content
    text = message.get("content")
    reasoning = message.get("reasoning")
    if text:
        content_blocks.append({"type": "text", "text": text})
    elif reasoning:
        content_blocks.append({"type": "text", "text": reasoning})

    # Tool calls
    tool_calls = message.get("tool_calls", [])
    for tc in tool_calls:
        func = tc.get("function", {})
        raw_args = func.get("arguments") or ""
        raw_args_str = raw_args.strip() if isinstance(raw_args, str) else ""

        if not raw_args_str:
            inp = {}
        else:
            try:
                inp = json.loads(raw_args_str)
                if not isinstance(inp, dict):
                    inp = {"value": inp}
            except json.JSONDecodeError:
                inp = {"raw": raw_args_str}

        content_blocks.append({
            "type": "tool_use",
            "id": tc.get("id", f"toolu_{uuid.uuid4().hex[:24]}"),
            "name": func.get("name", ""),
            "input": inp,
        })

    if not content_blocks:
        content_blocks.append({"type": "text", "text": ""})

    # Usage
    usage_in = openai_resp.get("usage", {})
    usage = {
        "input_tokens": usage_in.get("prompt_tokens", 0),
        "output_tokens": usage_in.get("completion_tokens", 0),
    }

    return {
        "id": f"msg_{uuid.uuid4().hex[:24]}",
        "type": "message",
        "role": "assistant",
        "content": content_blocks,
        "model": model_alias,
        "stop_reason": FINISH_REASON_MAP.get(finish, "end_turn"),
        "stop_sequence": None,
        "usage": usage,
    }


# =============================================================================
# PROTOCOL TRANSLATION: OpenAI SSE → Anthropic SSE (Streaming)
# =============================================================================

class StreamTranslator:
    """
    Stateful translator that converts OpenAI streaming chunks into
    strictly compliant Anthropic SSE events.
    """

    def __init__(self, wfile, model_alias: str):
        self.wfile = wfile
        self.model_alias = model_alias
        self.msg_id = f"msg_{uuid.uuid4().hex[:24]}"
        self.started = False
        self._finished = False

        # Current open content block tracking (Anthropic requires strict 1-block-at-a-time)
        self.current_block_type: Optional[str] = None  # "text" or "tool_use"
        self.current_block_index: int = -1
        self.next_block_index: int = 0

        # Tool tracking: openai tc_index → state
        self.tool_blocks: Dict[int, Dict] = {}

        self.input_tokens = 0
        self.output_tokens = 0
        self.accumulated_text: List[str] = []
        self.finish_reason = "end_turn"
        self.last_usage: Optional[Dict] = None

    def _write_event(self, event: str, data: Any) -> bool:
        """Write a single SSE event to the response stream. Returns False on client disconnect."""
        try:
            payload = f"event: {event}\ndata: {json.dumps(data)}\n\n"
            self.wfile.write(payload.encode("utf-8"))
            self.wfile.flush()
            return True
        except (BrokenPipeError, ConnectionResetError, OSError):
            return False

    def emit_message_start(self):
        """Emit the initial message_start event."""
        if self.started:
            return
        self.started = True
        self._write_event("message_start", {
            "type": "message_start",
            "message": {
                "id": self.msg_id,
                "type": "message",
                "role": "assistant",
                "content": [],
                "model": self.model_alias,
                "stop_reason": None,
                "stop_sequence": None,
                "usage": {"input_tokens": self.input_tokens, "output_tokens": 0},
            },
        })
        self._write_event("ping", {"type": "ping"})

    def _close_current_block(self):
        """Close the currently open content block if any."""
        if self.current_block_type is not None and self.current_block_index >= 0:
            self._write_event("content_block_stop", {
                "type": "content_block_stop",
                "index": self.current_block_index,
            })
            self.current_block_type = None
            self.current_block_index = -1

    def emit_text_delta(self, text: str):
        """Emit a text delta, opening a text block if needed."""
        self.emit_message_start()

        if self.current_block_type != "text":
            self._close_current_block()
            self.current_block_type = "text"
            self.current_block_index = self.next_block_index
            self.next_block_index += 1
            self._write_event("content_block_start", {
                "type": "content_block_start",
                "index": self.current_block_index,
                "content_block": {"type": "text", "text": ""},
            })

        self.accumulated_text.append(text)
        self._write_event("content_block_delta", {
            "type": "content_block_delta",
            "index": self.current_block_index,
            "delta": {"type": "text_delta", "text": text},
        })

    def handle_tool_call_delta(self, tc_index: int, tc_delta: Dict):
        """Handle a tool call chunk from OpenAI streaming."""
        self.emit_message_start()

        func = tc_delta.get("function", {})
        args = func.get("arguments", "")
        tc_id = tc_delta.get("id")
        name = func.get("name", "")

        if tc_index not in self.tool_blocks:
            # Starting a new tool call — close whatever block was open
            self._close_current_block()

            block_idx = self.next_block_index
            self.next_block_index += 1
            tool_id = tc_id or f"toolu_{uuid.uuid4().hex[:24]}"
            tool_name = name or "unknown"

            self.tool_blocks[tc_index] = {
                "block_index": block_idx,
                "id": tool_id,
                "name": tool_name,
            }
            self.current_block_type = "tool_use"
            self.current_block_index = block_idx

            self._write_event("content_block_start", {
                "type": "content_block_start",
                "index": block_idx,
                "content_block": {
                    "type": "tool_use",
                    "id": tool_id,
                    "name": tool_name,
                    "input": {},
                },
            })
            if args:
                self._write_event("content_block_delta", {
                    "type": "content_block_delta",
                    "index": block_idx,
                    "delta": {"type": "input_json_delta", "partial_json": args},
                })
        else:
            # Continuation of existing tool call
            state = self.tool_blocks[tc_index]
            block_idx = state["block_index"]

            # If current active block isn't this tool, switch back
            if self.current_block_index != block_idx:
                self._close_current_block()
                self.current_block_type = "tool_use"
                self.current_block_index = block_idx

            if args:
                self._write_event("content_block_delta", {
                    "type": "content_block_delta",
                    "index": block_idx,
                    "delta": {"type": "input_json_delta", "partial_json": args},
                })

    def finish(self, usage_data: Optional[Dict] = None):
        """Emit the message_delta and message_stop events."""
        if self._finished:
            return
        self._finished = True

        if not self.started:
            self.emit_message_start()
            self.emit_text_delta("")

        self._close_current_block()

        # Token counts
        u = usage_data or self.last_usage
        if u:
            self.output_tokens = u.get("completion_tokens", 0)
            self.input_tokens = u.get("prompt_tokens", self.input_tokens)

        # Fallback estimation if backend didn't supply token counts
        if self.output_tokens == 0:
            total_chars = sum(len(t) for t in self.accumulated_text)
            self.output_tokens = max(1, total_chars // 4)

        self._write_event("message_delta", {
            "type": "message_delta",
            "delta": {
                "stop_reason": self.finish_reason,
                "stop_sequence": None,
            },
            "usage": {"output_tokens": self.output_tokens},
        })
        self._write_event("message_stop", {"type": "message_stop"})

    def process_chunk(self, chunk: Dict) -> str:
        """Process a single OpenAI streaming chunk."""
        # Capture usage if present in any chunk
        if "usage" in chunk and chunk["usage"]:
            self.last_usage = chunk["usage"]

        choices = chunk.get("choices", [])

        # Usage-only chunk (e.g. final chunk with stream_options.include_usage)
        if not choices and "usage" in chunk:
            self.finish(chunk.get("usage"))
            return "done"

        if not choices:
            return "continue"

        choice = choices[0]
        delta = choice.get("delta", {})
        finish_reason = choice.get("finish_reason")

        # Text / reasoning content
        text = delta.get("content") or delta.get("reasoning")
        if text:
            self.emit_text_delta(text)

        # Tool calls
        tool_calls = delta.get("tool_calls", [])
        for tc in tool_calls:
            tc_idx = tc.get("index", 0)
            self.handle_tool_call_delta(tc_idx, tc)

        # Finish reason
        if finish_reason:
            self.finish_reason = FINISH_REASON_MAP.get(finish_reason, "end_turn")

        return "continue"


# =============================================================================
# ANTHROPIC ERROR RESPONSE BUILDER
# =============================================================================

ERROR_TYPE_MAP = {
    400: "invalid_request_error",
    401: "authentication_error",
    403: "permission_error",
    404: "not_found_error",
    429: "rate_limit_error",
    500: "api_error",
    502: "api_error",
    503: "overloaded_error",
    529: "overloaded_error",
}


def build_error_response(status: int, message: str) -> bytes:
    """Build an Anthropic-format error JSON response."""
    return json.dumps({
        "type": "error",
        "error": {
            "type": ERROR_TYPE_MAP.get(status, "api_error"),
            "message": message,
        },
    }).encode("utf-8")


# =============================================================================
# GATEWAY CONFIGURATION (Runtime State)
# =============================================================================

class GatewayConfig:
    """Holds the active gateway configuration set via CLI."""

    def __init__(self):
        self.provider_name: str = ""
        self.base_url: str = ""
        self.api_key: str = ""
        self.backend_model: str = ""
        self.host: str = DEFAULT_HOST
        self.port: int = DEFAULT_PORT
        self.insecure_ssl: bool = False
        self.request_count: int = 0
        self.start_time: float = 0.0
        self.lock = threading.Lock()

    def log_request(self, method: str, path: str, model: str, status: int,
                    elapsed_ms: float, in_tokens: int, out_tokens: int):
        """Log a request to the console."""
        with self.lock:
            self.request_count += 1

        ts = datetime.now().strftime("%H:%M:%S")
        status_color = C.GREEN if 200 <= status < 300 else C.RED
        print(
            f"  {C.DIM}{ts}{C.RESET} "
            f"{C.CYAN}{method}{C.RESET} {path} "
            f"-> {C.BOLD}{model}{C.RESET} "
            f"{status_color}{status}{C.RESET} "
            f"{C.DIM}({elapsed_ms:.0f}ms, {in_tokens}in/{out_tokens}out){C.RESET}"
        )


# Global config instance
CONFIG = GatewayConfig()


# =============================================================================
# HTTP GATEWAY HANDLER
# =============================================================================

class GatewayHandler(BaseHTTPRequestHandler):
    """Handles incoming Anthropic API requests and proxies to OpenAI backend."""

    # Suppress default access logging
    def log_message(self, format, *args):
        pass

    def _normalize_path(self) -> str:
        """Normalize URL path, stripping query parameters and repeated /v1."""
        raw_path = self.path.split("?")[0].rstrip("/")
        parts = [p for p in raw_path.split("/") if p]
        # Collapse repeated v1 (e.g. /v1/v1/messages -> /v1/messages)
        if len(parts) >= 2 and parts[0] == "v1" and parts[1] == "v1":
            parts.pop(0)
        return "/" + "/".join(parts)

    def _send_json(self, status: int, body: bytes, headers: Optional[Dict] = None):
        """Send a JSON response."""
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("anthropic-version", "2023-06-01")
        self.send_header("x-request-id", str(uuid.uuid4()))
        self.send_header("Access-Control-Allow-Origin", "*")
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _send_error(self, status: int, message: str, headers: Optional[Dict] = None):
        """Send an Anthropic-format error response."""
        self._send_json(status, build_error_response(status, message), headers)

    def _start_sse(self):
        """Begin an SSE streaming response."""
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "keep-alive")
        self.send_header("anthropic-version", "2023-06-01")
        self.send_header("x-request-id", str(uuid.uuid4()))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def do_OPTIONS(self):
        """CORS preflight handling."""
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    # ---- Routes ----

    def do_GET(self):
        path = self._normalize_path()

        if path in ("/healthz", "/health"):
            self._send_json(200, json.dumps({
                "status": "ok",
                "version": VERSION,
                "provider": CONFIG.provider_name,
                "model": CONFIG.backend_model,
                "uptime_s": int(time.time() - CONFIG.start_time),
                "requests": CONFIG.request_count,
            }).encode())
            return

        if path in ("/v1/models", "/models"):
            # Return model aliases for Claude Code compatibility
            models = [
                {"id": CONFIG.backend_model, "object": "model", "display_name": CONFIG.backend_model},
                {"id": "claude-sonnet-4-20250514", "object": "model", "display_name": "Claude Sonnet 4"},
                {"id": "claude-3-5-sonnet-20241022", "object": "model", "display_name": "Claude 3.5 Sonnet"},
                {"id": "claude-3-5-haiku-20241022", "object": "model", "display_name": "Claude 3.5 Haiku"},
                {"id": "claude-3-haiku-20240307", "object": "model", "display_name": "Claude 3 Haiku"},
                {"id": "claude-3-opus-20240229", "object": "model", "display_name": "Claude 3 Opus"},
            ]
            self._send_json(200, json.dumps({"object": "list", "data": models}).encode())
            return

        self._send_error(404, f"Not found: {self.path}")

    def do_POST(self):
        path = self._normalize_path()

        if path in ("/v1/messages", "/messages"):
            self._handle_messages()
            return

        if path in ("/v1/messages/count_tokens", "/messages/count_tokens"):
            self._handle_count_tokens()
            return

        self._send_error(404, f"Not found: {self.path}")

    # ---- /v1/messages handler ----

    def _handle_messages(self):
        t0 = time.time()

        # Read request body
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            body = json.loads(raw)
        except Exception as e:
            self._send_error(400, f"Invalid request body: {e}")
            return

        model_alias = body.get("model", "claude-sonnet-4-20250514")
        is_stream = body.get("stream", False)

        # Translate to OpenAI format
        try:
            openai_body = build_openai_request(body, CONFIG.backend_model)
        except Exception as e:
            self._send_error(400, f"Request translation error: {e}")
            return

        # Build upstream request
        url = build_chat_url(CONFIG.base_url)
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream" if is_stream else "application/json",
            "Authorization": f"Bearer {CONFIG.api_key}",
            "User-Agent": f"ccgate/{VERSION}",
        }
        # OpenRouter-specific headers
        if "openrouter.ai" in url:
            headers["HTTP-Referer"] = "https://github.com/Key-Switcher/ccgate"
            headers["X-Title"] = "ccgate - Claude Code Gateway"

        post_data = json.dumps(openai_body).encode("utf-8")
        req = urllib.request.Request(url, data=post_data, headers=headers, method="POST")

        ctx = ssl.create_default_context()
        if CONFIG.insecure_ssl or os.environ.get("GATEWAY_INSECURE_SSL") == "1":
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

        if is_stream:
            self._handle_stream(req, ctx, model_alias, t0)
        else:
            self._handle_sync(req, ctx, model_alias, t0)

    def _handle_sync(self, req, ctx, model_alias: str, t0: float):
        """Handle a non-streaming /v1/messages request."""
        try:
            with urllib.request.urlopen(req, timeout=180, context=ctx) as resp:
                body = json.loads(resp.read().decode("utf-8", errors="replace"))
        except urllib.error.HTTPError as e:
            status = e.code
            try:
                err_body = json.loads(e.read().decode("utf-8", errors="replace"))
                msg = err_body.get("error", {}).get("message", str(e))
            except Exception:
                msg = str(e)
            elapsed = (time.time() - t0) * 1000
            CONFIG.log_request("POST", "/v1/messages", CONFIG.backend_model, status, elapsed, 0, 0)
            self._send_error(status, msg)
            return
        except Exception as e:
            elapsed = (time.time() - t0) * 1000
            CONFIG.log_request("POST", "/v1/messages", CONFIG.backend_model, 502, elapsed, 0, 0)
            self._send_error(502, f"Upstream error: {e}")
            return

        # Translate response
        translated = translate_openai_response(body, model_alias)
        resp_bytes = json.dumps(translated).encode("utf-8")

        in_tok = translated.get("usage", {}).get("input_tokens", 0)
        out_tok = translated.get("usage", {}).get("output_tokens", 0)
        elapsed = (time.time() - t0) * 1000
        CONFIG.log_request("POST", "/v1/messages", CONFIG.backend_model, 200, elapsed, in_tok, out_tok)

        self._send_json(200, resp_bytes)

    def _handle_stream(self, req, ctx, model_alias: str, t0: float):
        """Handle a streaming /v1/messages request with SSE translation."""
        try:
            resp = urllib.request.urlopen(req, timeout=180, context=ctx)
        except urllib.error.HTTPError as e:
            status = e.code
            try:
                err_body = json.loads(e.read().decode("utf-8", errors="replace"))
                msg = err_body.get("error", {}).get("message", str(e))
            except Exception:
                msg = str(e)
            elapsed = (time.time() - t0) * 1000
            CONFIG.log_request("POST", "/v1/messages (stream)", CONFIG.backend_model, status, elapsed, 0, 0)

            retry_after = e.headers.get("retry-after") if e.headers else None
            hdrs = {"retry-after": retry_after} if retry_after else {}
            self._send_error(status, msg, hdrs)
            return
        except Exception as e:
            elapsed = (time.time() - t0) * 1000
            CONFIG.log_request("POST", "/v1/messages (stream)", CONFIG.backend_model, 502, elapsed, 0, 0)
            self._send_error(502, f"Upstream connection error: {e}")
            return

        # Start SSE response to Claude Code
        self._start_sse()
        translator = StreamTranslator(self.wfile, model_alias)
        finished = False

        try:
            for line_bytes in resp:
                line = line_bytes.decode("utf-8", errors="replace").strip()
                if not line or line.startswith(":"):
                    continue

                if line.startswith("data:"):
                    data_str = line[5:].strip()
                    if data_str == "[DONE]":
                        if not finished:
                            translator.finish()
                            finished = True
                        break
                    try:
                        chunk = json.loads(data_str)
                        result = translator.process_chunk(chunk)
                        if result == "done":
                            finished = True
                    except json.JSONDecodeError:
                        continue
        except (BrokenPipeError, ConnectionResetError):
            # Client disconnected (e.g. user pressed Ctrl+C in Claude Code)
            finished = True
        except Exception as e:
            if not finished:
                try:
                    translator._write_event("error", {
                        "type": "error",
                        "error": {
                            "type": "api_error",
                            "message": f"Stream interrupted: {e}",
                        },
                    })
                except Exception:
                    pass
        finally:
            try:
                resp.close()
            except Exception:
                pass

        if not finished:
            translator.finish()

        self.close_connection = True

        elapsed = (time.time() - t0) * 1000
        CONFIG.log_request(
            "POST", "/v1/messages (stream)", CONFIG.backend_model, 200,
            elapsed, translator.input_tokens, translator.output_tokens,
        )

    def _handle_count_tokens(self):
        """Proxy count_tokens — returns a token estimate."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            body = json.loads(raw)
        except Exception as e:
            self._send_error(400, f"Invalid request body: {e}")
            return

        total_chars = 0
        messages = body.get("messages", [])
        for msg in messages:
            content = msg.get("content", "")
            if isinstance(content, str):
                total_chars += len(content)
            elif isinstance(content, list):
                for block in content:
                    if isinstance(block, dict):
                        total_chars += len(block.get("text", ""))

        system = body.get("system", "")
        if isinstance(system, str):
            total_chars += len(system)
        elif isinstance(system, list):
            for b in system:
                if isinstance(b, dict):
                    total_chars += len(b.get("text", ""))

        tools = body.get("tools", [])
        for tool in tools:
            if isinstance(tool, dict):
                total_chars += len(tool.get("name", "")) + len(tool.get("description", ""))
                total_chars += len(json.dumps(tool.get("input_schema", {})))

        est_tokens = max(1, total_chars // 4)
        self._send_json(200, json.dumps({"input_tokens": est_tokens}).encode())


# =============================================================================
# THREADED HTTP SERVER
# =============================================================================

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Thread-per-request HTTP server."""
    daemon_threads = True
    allow_reuse_address = True


# =============================================================================
# CLI — INTERACTIVE PROVIDER / MODEL / KEY SELECTION
# =============================================================================

def print_banner():
    """Print the startup banner."""
    print(f"""
{C.BLUE}{C.BOLD}+----------------------------------------------------------+
|           ccgate - Claude Code Gateway  v{VERSION}           |
+----------------------------------------------------------+{C.RESET}
{C.DIM}Routes Claude Code through OpenAI-compatible providers{C.RESET}
""")


def select_from_list(prompt: str, items: List[str], allow_custom: bool = False) -> int:
    """Display a numbered list and get user selection."""
    for i, item in enumerate(items, 1):
        print(f"  {C.CYAN}[{i}]{C.RESET} {item}")
    if allow_custom:
        print(f"  {C.CYAN}[{len(items) + 1}]{C.RESET} {C.DIM}Enter manually{C.RESET}")

    while True:
        try:
            max_val = len(items) + (1 if allow_custom else 0)
            choice = input(f"\n{prompt} [{C.BOLD}1-{max_val}{C.RESET}]: ").strip()
            idx = int(choice) - 1
            if 0 <= idx < max_val:
                return idx
        except (ValueError, EOFError):
            pass
        print(f"  {C.RED}Invalid choice. Try again.{C.RESET}")


def interactive_setup() -> bool:
    """Interactive CLI to configure the gateway. Returns True if config is set."""
    providers = load_providers_json()
    saved_keys = load_keys_from_excel()

    if not providers:
        print(f"{C.RED}No providers found in {PROVIDERS_FILE}{C.RESET}")
        print(f"{C.DIM}Add providers with api_tester.py or edit providers.json manually.{C.RESET}")
        return False

    # 1. Select provider
    print(f"{C.BOLD}Select backend provider:{C.RESET}")
    prov_names = list(providers.keys())
    prov_items = []
    for name in prov_names:
        url = providers[name].get("base_url", "")
        model_count = len(providers[name].get("models", []))
        prov_items.append(f"{name:<16} {C.DIM}({url}, {model_count} models){C.RESET}")

    idx = select_from_list("Provider", prov_items)
    prov_name = prov_names[idx]
    prov_info = providers[prov_name]

    CONFIG.provider_name = prov_name
    CONFIG.base_url = prov_info.get("base_url", "")

    # 2. Select model
    models = prov_info.get("models", [])
    if not models:
        print(f"\n{C.YELLOW}No models listed for {prov_name}. Enter model name:{C.RESET}")
        CONFIG.backend_model = input("  Model: ").strip()
        if not CONFIG.backend_model:
            print(f"{C.RED}Model is required.{C.RESET}")
            return False
    else:
        print(f"\n{C.BOLD}Select model:{C.RESET}")
        idx = select_from_list("Model", models, allow_custom=True)
        if idx < len(models):
            CONFIG.backend_model = models[idx]
        else:
            CONFIG.backend_model = input("  Enter model name: ").strip()
            if not CONFIG.backend_model:
                print(f"{C.RED}Model is required.{C.RESET}")
                return False

    # 3. Select API key
    matching_keys = {}
    p_low = prov_name.strip().lower()
    for nick, data in saved_keys.items():
        prov = data.get("provider", "").strip().lower()
        if prov == p_low or p_low in prov or p_low in nick.lower():
            matching_keys[nick] = data

    if matching_keys:
        print(f"\n{C.BOLD}Select API key:{C.RESET}")
        key_names = list(matching_keys.keys())
        key_items = [f"{nick} {C.DIM}({matching_keys[nick]['provider']}){C.RESET}" for nick in key_names]
        idx = select_from_list("API Key", key_items, allow_custom=True)
        if idx < len(key_names):
            k = matching_keys[key_names[idx]]["key"]
            if k.lower().startswith("bearer "):
                k = k[7:].strip()
            CONFIG.api_key = k
        else:
            CONFIG.api_key = input("  Enter API key: ").strip()
    else:
        all_key_names = list(saved_keys.keys())
        if all_key_names:
            print(f"\n{C.BOLD}Select API key:{C.RESET}")
            key_items = [f"{n} {C.DIM}({saved_keys[n]['provider']}){C.RESET}" for n in all_key_names]
            idx = select_from_list("API Key", key_items, allow_custom=True)
            if idx < len(all_key_names):
                k = saved_keys[all_key_names[idx]]["key"]
                if k.lower().startswith("bearer "):
                    k = k[7:].strip()
                CONFIG.api_key = k
            else:
                CONFIG.api_key = input("  Enter API key: ").strip()
        else:
            print(f"\n{C.YELLOW}No saved keys found. Enter API key manually:{C.RESET}")
            CONFIG.api_key = input("  API Key: ").strip()

    if not CONFIG.api_key:
        print(f"{C.RED}API key is required.{C.RESET}")
        return False

    return True


# =============================================================================
# COMMANDS: setup, doctor
# =============================================================================

def cmd_setup(model_to_set: Optional[str] = None):
    """Write gateway config to Claude Code's settings.json."""
    url = f"http://{CONFIG.host}:{CONFIG.port}"
    print(f"\n{C.BOLD}Configuring Claude Code...{C.RESET}")

    settings = {}
    if CLAUDE_SETTINGS.exists():
        try:
            settings = json.loads(CLAUDE_SETTINGS.read_text(encoding="utf-8"))
        except Exception:
            settings = {}

    env = settings.get("env", {})
    env["ANTHROPIC_BASE_URL"] = url
    env["ANTHROPIC_AUTH_TOKEN"] = "ccgate-local"

    chosen_model = model_to_set or CONFIG.backend_model
    if chosen_model:
        env["ANTHROPIC_MODEL"] = chosen_model
        env["ANTHROPIC_SMALL_FAST_MODEL"] = chosen_model
        env["ANTHROPIC_DEFAULT_SONNET_MODEL"] = chosen_model
        env["ANTHROPIC_DEFAULT_OPUS_MODEL"] = chosen_model
        env["ANTHROPIC_DEFAULT_HAIKU_MODEL"] = chosen_model
        settings["model"] = chosen_model

    settings["env"] = env

    CLAUDE_SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    CLAUDE_SETTINGS.write_text(json.dumps(settings, indent=2) + "\n", encoding="utf-8")

    print(f"  {C.GREEN}[OK]{C.RESET} Updated {C.CYAN}{CLAUDE_SETTINGS}{C.RESET}")
    print(f"  {C.GREEN}[OK]{C.RESET} ANTHROPIC_BASE_URL = {C.BOLD}{url}{C.RESET}")
    print(f"  {C.GREEN}[OK]{C.RESET} ANTHROPIC_AUTH_TOKEN = {C.BOLD}ccgate-local{C.RESET}")
    if chosen_model:
        print(f"  {C.GREEN}[OK]{C.RESET} Model set to: {C.BOLD}{chosen_model}{C.RESET}")
    print(f"\n{C.YELLOW}Restart Claude Code for changes to take effect.{C.RESET}\n")


def cmd_doctor():
    """Test all configured providers with real HTTP calls."""
    print_banner()
    providers = load_providers_json()
    saved_keys = load_keys_from_excel()

    if not providers:
        print(f"{C.RED}No providers found in providers.json{C.RESET}")
        return

    print(f"{C.BOLD}Testing upstream providers...{C.RESET}\n")

    for prov_name, prov_info in providers.items():
        base_url = prov_info.get("base_url", "")
        models = prov_info.get("models", [])
        model = models[0] if models else "test-model"

        api_key = find_key_for_provider(prov_name, saved_keys)
        if api_key.lower().startswith("bearer "):
            api_key = api_key[7:].strip()

        if not api_key:
            print(f"  {C.YELLOW}[?]{C.RESET} {prov_name:<16} -- {C.DIM}No API key found in keys.xlsx, skipping{C.RESET}")
            continue

        url = build_chat_url(base_url)
        payload = json.dumps({
            "model": model,
            "messages": [{"role": "user", "content": "Say OK"}],
            "max_tokens": 5,
        }).encode("utf-8")

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
            "User-Agent": f"ccgate/{VERSION}",
        }
        if "openrouter.ai" in url:
            headers["HTTP-Referer"] = "https://github.com/Key-Switcher/ccgate"
            headers["X-Title"] = "ccgate - Claude Code Gateway"

        req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
        ctx = ssl.create_default_context()
        if CONFIG.insecure_ssl or os.environ.get("GATEWAY_INSECURE_SSL") == "1":
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE

        try:
            t0 = time.time()
            with urllib.request.urlopen(req, timeout=20, context=ctx) as resp:
                elapsed = (time.time() - t0) * 1000
                body = json.loads(resp.read().decode("utf-8", errors="replace"))
                reply = ""
                if body.get("choices"):
                    msg = body["choices"][0].get("message") or {}
                    reply = (msg.get("content") or msg.get("reasoning") or "")[:50].strip().replace("\n", " ")
                print(
                    f"  {C.GREEN}[OK]{C.RESET} {prov_name:<16} "
                    f"{C.DIM}({model}){C.RESET} -- "
                    f"{C.GREEN}{elapsed:.0f}ms{C.RESET} "
                    f"{C.DIM}\"{reply}\"{C.RESET}"
                )
        except urllib.error.HTTPError as e:
            elapsed = (time.time() - t0) * 1000
            try:
                err_detail = json.loads(e.read().decode("utf-8", errors="replace"))
                msg = err_detail.get("error", {}).get("message", e.reason)
            except Exception:
                msg = e.reason
            print(
                f"  {C.RED}[FAIL]{C.RESET} {prov_name:<16} "
                f"{C.DIM}({model}){C.RESET} -- "
                f"{C.RED}{e.code} {msg[:60]}{C.RESET} "
                f"{C.DIM}({elapsed:.0f}ms){C.RESET}"
            )
        except Exception as e:
            print(
                f"  {C.RED}[FAIL]{C.RESET} {prov_name:<16} "
                f"{C.DIM}({model}){C.RESET} -- "
                f"{C.RED}{str(e)[:60]}{C.RESET}"
            )

    print()


# =============================================================================
# SERVER START
# =============================================================================

def start_server():
    """Start the gateway HTTP server."""
    CONFIG.start_time = time.time()

    server = ThreadedHTTPServer((CONFIG.host, CONFIG.port), GatewayHandler)

    # Graceful shutdown on Ctrl+C
    def shutdown_handler(signum, frame):
        print(f"\n\n{C.YELLOW}Shutting down gateway...{C.RESET}")
        threading.Thread(target=server.shutdown).start()

    signal.signal(signal.SIGINT, shutdown_handler)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, shutdown_handler)

    url = f"http://{CONFIG.host}:{CONFIG.port}"

    print(f"""
{C.GREEN}{C.BOLD}[OK] Gateway running{C.RESET} on {C.CYAN}{C.BOLD}{url}{C.RESET}

{C.BOLD}  Provider :{C.RESET} {CONFIG.provider_name}
{C.BOLD}  Model    :{C.RESET} {CONFIG.backend_model}
{C.BOLD}  Endpoint :{C.RESET} {build_chat_url(CONFIG.base_url)}

{C.BOLD}Claude Code configuration:{C.RESET}
  {C.CYAN}ANTHROPIC_BASE_URL{C.RESET}={C.BOLD}{url}{C.RESET}
  {C.CYAN}ANTHROPIC_AUTH_TOKEN{C.RESET}={C.BOLD}ccgate-local{C.RESET}

{C.DIM}Press Ctrl+C to stop.{C.RESET}
{"-" * 60}""")

    try:
        server.serve_forever()
    except Exception:
        pass
    finally:
        server.server_close()
        print(f"{C.GREEN}Gateway stopped.{C.RESET}\n")


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="ccgate - Claude Code Gateway",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Commands:
  (default)    Interactive provider/model selection, then start gateway
  setup        Configure Claude Code to use the gateway
  doctor       Test all provider connections

Examples:
  python ccgate.py                                              # Interactive
  python ccgate.py --provider groq --model openai/gpt-oss-20b  # Direct start
  python ccgate.py --provider groq --setup                     # Configure Claude Code and run
  python ccgate.py setup                                       # Configure Claude settings
  python ccgate.py doctor                                      # Health check
""",
    )
    parser.add_argument("command", nargs="?", default=None, help="Command: setup, doctor")
    parser.add_argument("--provider", "-p", help="Provider name (from providers.json)")
    parser.add_argument("--model", "-m", help="Backend model name")
    parser.add_argument("--key", "-k", help="API key (or auto-found from keys.xlsx)")
    parser.add_argument("--key-name", help="Nickname from keys.xlsx")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT, help=f"Port (default: {DEFAULT_PORT})")
    parser.add_argument("--host", default=DEFAULT_HOST, help=f"Host (default: {DEFAULT_HOST})")
    parser.add_argument("--base-url", help="Override base URL for the provider")
    parser.add_argument("--setup", action="store_true", help="Automatically write ~/.claude/settings.json before starting")
    parser.add_argument("--insecure", action="store_true", help="Disable SSL certificate verification")

    args = parser.parse_args()
    CONFIG.port = args.port
    CONFIG.host = args.host
    CONFIG.insecure_ssl = args.insecure

    # --- Standalone Commands ---
    if args.command == "doctor":
        cmd_doctor()
        return

    if args.command == "setup":
        print_banner()
        cmd_setup(args.model)
        return

    print_banner()

    # --- CLI args mode ---
    if args.provider:
        providers = load_providers_json()
        saved_keys = load_keys_from_excel()

        # Find provider (case-insensitive)
        matched = None
        for pname in providers:
            if pname.strip().lower() == args.provider.strip().lower():
                matched = pname
                break
        if not matched:
            print(f"{C.RED}Provider '{args.provider}' not found in providers.json{C.RESET}")
            print(f"{C.DIM}Available: {', '.join(providers.keys())}{C.RESET}")
            sys.exit(1)

        CONFIG.provider_name = matched
        CONFIG.base_url = args.base_url or providers[matched].get("base_url", "")

        if args.model:
            CONFIG.backend_model = args.model
        else:
            models = providers[matched].get("models", [])
            if models:
                CONFIG.backend_model = models[0]
            else:
                print(f"{C.RED}No model specified and none listed for {matched}{C.RESET}")
                sys.exit(1)

        if args.key:
            k = args.key.strip()
            CONFIG.api_key = k[7:].strip() if k.lower().startswith("bearer ") else k
        elif args.key_name:
            if args.key_name in saved_keys:
                k = saved_keys[args.key_name]["key"]
                CONFIG.api_key = k[7:].strip() if k.lower().startswith("bearer ") else k
            else:
                print(f"{C.RED}Key '{args.key_name}' not found in keys.xlsx{C.RESET}")
                sys.exit(1)
        else:
            CONFIG.api_key = find_key_for_provider(matched, saved_keys)
            if not CONFIG.api_key:
                print(f"{C.RED}No API key found for {matched}. Use --key or --key-name.{C.RESET}")
                sys.exit(1)

    # --- Interactive mode ---
    else:
        if not interactive_setup():
            sys.exit(1)

    # Automatically configure Claude Code if requested
    if args.setup:
        cmd_setup()

    # --- Start ---
    start_server()


if __name__ == "__main__":
    main()

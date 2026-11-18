"""
===============================================================================
ccgate.py — Claude Code Gateway (GUI & CLI)
===============================================================================
A self-hosted protocol-translating reverse proxy that lets Claude Code talk to
ANY OpenAI-compatible backend provider (OpenRouter, Groq, NVIDIA AI Studio,
Hugging Face, Ollama, etc.).

Features:
    • Modern PySide6 GUI control panel with live traffic monitor
    • Full CLI support with --provider, doctor, setup, and --cli flags
    • Translates Anthropic Messages API <-> OpenAI Chat Completions API
    • Full SSE streaming with real-time event translation
    • Tool call / function call bidirectional translation (stream & sync)
    • Automatic port reclaim & cleanup (no more locked port 5005 on restart)
    • Integrates with providers.json and keys.xlsx for quick selection
    • 1-Click Claude Code configuration (~/.claude/settings.json)
===============================================================================
"""

import sys
import os
import json
import time
import uuid
import ssl
import socket
import argparse
import signal
import atexit
import threading
import subprocess
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

VERSION = "1.2.0"
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

    @staticmethod
    def disable():
        for attr in dir(C):
            if attr.isupper() and not attr.startswith("_"):
                setattr(C, attr, "")


if not sys.stdout.isatty():
    C.disable()


# =============================================================================
# PORT UTILITIES & CLEANUP (Solves Port 5005 Re-use on Windows)
# =============================================================================

def is_port_in_use(port: int, host: str = DEFAULT_HOST) -> bool:
    """Check if a TCP port is currently listening."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex((host, port)) == 0


def free_port_if_in_use(port: int, host: str = DEFAULT_HOST) -> bool:
    """
    Cleanly reclaims a port if an orphaned ccgate or Python process is still holding it.
    1. Sends graceful /_shutdown HTTP probe
    2. If still busy, finds and terminates the stale process on Windows
    """
    if not is_port_in_use(port, host):
        return True

    # 1. Try graceful shutdown probe
    try:
        req = urllib.request.Request(f"http://{host}:{port}/_shutdown")
        with urllib.request.urlopen(req, timeout=1):
            pass
        for _ in range(5):
            time.sleep(0.2)
            if not is_port_in_use(port, host):
                return True
    except Exception:
        pass

    # 2. Terminate stale process holding port (Windows)
    if sys.platform == "win32":
        try:
            cmd = f'netstat -ano -p tcp | findstr ":{port} "'
            out = subprocess.check_output(cmd, shell=True, text=True, errors="replace")
            current_pid = os.getpid()
            for line in out.strip().splitlines():
                parts = line.split()
                if len(parts) >= 5 and "LISTENING" in parts[3].upper():
                    pid = int(parts[4])
                    if pid > 0 and pid != current_pid:
                        subprocess.call(f"taskkill /F /PID {pid}", shell=True,
                                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            for _ in range(5):
                time.sleep(0.2)
                if not is_port_in_use(port, host):
                    return True
        except Exception:
            pass

    return not is_port_in_use(port, host)


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

    for nick, data in saved_keys.items():
        if data.get("provider", "").strip().lower() == p_low:
            return data.get("key", "")

    for nick, data in saved_keys.items():
        prov = data.get("provider", "").strip().lower()
        if p_low in prov or prov in p_low:
            return data.get("key", "")

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
# PROTOCOL TRANSLATION: Anthropic Request -> OpenAI Request
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

            for tr in tool_results:
                result.append(tr)

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

    # 1. System prompt -> system message
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

    tools = anthropic_body.get("tools")
    if tools:
        openai_req["tools"] = translate_tools_to_openai(tools)
    tc = anthropic_body.get("tool_choice")
    if tc:
        openai_req["tool_choice"] = translate_tool_choice_to_openai(tc)

    return openai_req


# =============================================================================
# PROTOCOL TRANSLATION: OpenAI Response -> Anthropic Response (Non-Streaming)
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

    content_blocks = []

    text = message.get("content")
    reasoning = message.get("reasoning")
    if text:
        content_blocks.append({"type": "text", "text": text})
    elif reasoning:
        content_blocks.append({"type": "text", "text": reasoning})

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
# PROTOCOL TRANSLATION: OpenAI SSE -> Anthropic SSE (Streaming)
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

        self.current_block_type: Optional[str] = None
        self.current_block_index: int = -1
        self.next_block_index: int = 0

        self.tool_blocks: Dict[int, Dict] = {}

        self.input_tokens = 0
        self.output_tokens = 0
        self.accumulated_text: List[str] = []
        self.finish_reason = "end_turn"
        self.last_usage: Optional[Dict] = None

    def _write_event(self, event: str, data: Any) -> bool:
        """Write a single SSE event to the response stream."""
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
            state = self.tool_blocks[tc_index]
            block_idx = state["block_index"]

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

        u = usage_data or self.last_usage
        if u:
            self.output_tokens = u.get("completion_tokens", 0)
            self.input_tokens = u.get("prompt_tokens", self.input_tokens)

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
        if "usage" in chunk and chunk["usage"]:
            self.last_usage = chunk["usage"]

        choices = chunk.get("choices", [])
        if not choices and "usage" in chunk:
            self.finish(chunk.get("usage"))
            return "done"

        if not choices:
            return "continue"

        choice = choices[0]
        delta = choice.get("delta", {})
        finish_reason = choice.get("finish_reason")

        text = delta.get("content") or delta.get("reasoning")
        if text:
            self.emit_text_delta(text)

        tool_calls = delta.get("tool_calls", [])
        for tc in tool_calls:
            tc_idx = tc.get("index", 0)
            self.handle_tool_call_delta(tc_idx, tc)

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
# GATEWAY CONFIGURATION & RUNTIME STATE
# =============================================================================

class GatewayConfig:
    """Holds the active gateway configuration and callbacks."""

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
        self.is_running: bool = False
        self.server_instance: Optional['ThreadedHTTPServer'] = None
        self.server_thread: Optional[threading.Thread] = None
        self.log_callback = None
        self.lock = threading.Lock()

    def log_request(self, method: str, path: str, model: str, status: int,
                    elapsed_ms: float, in_tokens: int, out_tokens: int):
        """Log a request to the console and to GUI if registered."""
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

        if self.log_callback:
            try:
                self.log_callback(ts, method, path, model, status, elapsed_ms, in_tokens, out_tokens)
            except Exception:
                pass


CONFIG = GatewayConfig()


# =============================================================================
# HTTP GATEWAY HANDLER
# =============================================================================

class GatewayHandler(BaseHTTPRequestHandler):
    """Handles incoming Anthropic API requests and proxies to OpenAI backend."""

    def log_message(self, format, *args):
        pass

    def _normalize_path(self) -> str:
        raw_path = self.path.split("?")[0].rstrip("/")
        parts = [p for p in raw_path.split("/") if p]
        if len(parts) >= 2 and parts[0] == "v1" and parts[1] == "v1":
            parts.pop(0)
        return "/" + "/".join(parts)

    def _send_json(self, status: int, body: bytes, headers: Optional[Dict] = None):
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
        self._send_json(status, build_error_response(status, message), headers)

    def _start_sse(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "close")
        self.send_header("anthropic-version", "2023-06-01")
        self.send_header("x-request-id", str(uuid.uuid4()))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.end_headers()

    # ---- Routes ----

    def do_GET(self):
        path = self._normalize_path()

        if path in ("/_shutdown", "/shutdown"):
            self._send_json(200, b'{"status":"shutting_down"}')
            if CONFIG.server_instance:
                threading.Thread(target=CONFIG.server_instance.shutdown).start()
            return

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
            models = [
                {"id": CONFIG.backend_model, "object": "model", "display_name": CONFIG.backend_model},
                {"id": "claude-3-5-sonnet-20241022", "object": "model", "display_name": "Claude 3.5 Sonnet"},
                {"id": "claude-3-5-haiku-20241022", "object": "model", "display_name": "Claude 3.5 Haiku"},
                {"id": "claude-sonnet-4-20250514", "object": "model", "display_name": "Claude Sonnet 4"},
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

        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            body = json.loads(raw)
        except Exception as e:
            self._send_error(400, f"Invalid request body: {e}")
            return

        model_alias = body.get("model", "claude-3-5-sonnet-20241022")
        is_stream = body.get("stream", False)

        try:
            openai_body = build_openai_request(body, CONFIG.backend_model)
        except Exception as e:
            self._send_error(400, f"Request translation error: {e}")
            return

        url = build_chat_url(CONFIG.base_url)
        headers = {
            "Content-Type": "application/json",
            "Accept": "text/event-stream" if is_stream else "application/json",
            "Authorization": f"Bearer {CONFIG.api_key}",
            "User-Agent": f"ccgate/{VERSION}",
        }
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
# THREADED HTTP SERVER (WITH CLEAN SOCKET RE-USE)
# =============================================================================

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Thread-per-request HTTP server with clean address reuse."""
    daemon_threads = True
    allow_reuse_address = True

    def server_bind(self):
        try:
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        except Exception:
            pass
        super().server_bind()


# =============================================================================
# SERVER LIFECYCLE MANAGEMENT (START / STOP)
# =============================================================================

def start_server_thread(port: Optional[int] = None, host: Optional[str] = None) -> Tuple[bool, str]:
    """Start the gateway server in a background thread."""
    if CONFIG.is_running:
        return True, f"Gateway already running on http://{CONFIG.host}:{CONFIG.port}"

    if port:
        CONFIG.port = port
    if host:
        CONFIG.host = host

    # Reclaim port if left open by previous instance
    if not free_port_if_in_use(CONFIG.port, CONFIG.host):
        return False, f"Port {CONFIG.port} is busy and could not be released. Please try another port."

    CONFIG.start_time = time.time()
    CONFIG.request_count = 0

    try:
        server = ThreadedHTTPServer((CONFIG.host, CONFIG.port), GatewayHandler)
        CONFIG.server_instance = server
        CONFIG.is_running = True
    except Exception as e:
        CONFIG.is_running = False
        return False, f"Failed to bind to {CONFIG.host}:{CONFIG.port}: {e}"

    def serve():
        try:
            server.serve_forever()
        except Exception:
            pass
        finally:
            CONFIG.is_running = False
            try:
                server.server_close()
            except Exception:
                pass

    t = threading.Thread(target=serve, daemon=True)
    CONFIG.server_thread = t
    t.start()

    return True, f"Gateway started on http://{CONFIG.host}:{CONFIG.port}"


def stop_server_thread() -> Tuple[bool, str]:
    """Cleanly shut down the gateway server and release the port."""
    if not CONFIG.is_running or not CONFIG.server_instance:
        CONFIG.is_running = False
        return True, "Gateway is stopped."

    try:
        # Trigger shutdown from a separate thread so it doesn't deadlock
        threading.Thread(target=CONFIG.server_instance.shutdown).start()
    except Exception:
        pass

    try:
        CONFIG.server_instance.server_close()
    except Exception:
        pass

    CONFIG.is_running = False
    CONFIG.server_instance = None
    time.sleep(0.3)
    return True, "Gateway stopped cleanly."


# Register clean exit on process termination
atexit.register(stop_server_thread)


# =============================================================================
# CLAUDE CODE CONFIGURATION (settings.json)
# =============================================================================

def apply_claude_settings(port: int = DEFAULT_PORT, host: str = DEFAULT_HOST,
                          model_to_set: Optional[str] = None) -> Tuple[bool, str]:
    """Write gateway URL and model configuration to ~/.claude/settings.json."""
    url = f"http://{host}:{port}"
    try:
        settings = {}
        if CLAUDE_SETTINGS.exists():
            try:
                settings = json.loads(CLAUDE_SETTINGS.read_text(encoding="utf-8"))
            except Exception:
                settings = {}

        env = settings.get("env", {})
        env["ANTHROPIC_BASE_URL"] = url
        env["ANTHROPIC_AUTH_TOKEN"] = "ccgate-local"

        # Use standard Claude model alias so Claude Code sees a 200k context window without warnings
        chosen_alias = "claude-3-5-sonnet-20241022"
        env["ANTHROPIC_MODEL"] = chosen_alias
        env["ANTHROPIC_SMALL_FAST_MODEL"] = "claude-3-5-haiku-20241022"
        env["ANTHROPIC_DEFAULT_SONNET_MODEL"] = chosen_alias
        env["ANTHROPIC_DEFAULT_OPUS_MODEL"] = chosen_alias
        env["ANTHROPIC_DEFAULT_HAIKU_MODEL"] = "claude-3-5-haiku-20241022"
        settings["model"] = chosen_alias

        settings["env"] = env

        CLAUDE_SETTINGS.parent.mkdir(parents=True, exist_ok=True)
        CLAUDE_SETTINGS.write_text(json.dumps(settings, indent=2) + "\n", encoding="utf-8")
        return True, f"Claude Code configured! Base URL: {url}"
    except Exception as e:
        return False, f"Failed to update settings.json: {e}"


# =============================================================================
# DOCTOR (PROVIDER HEALTH TEST)
# =============================================================================

def test_single_provider(prov_name: str, base_url: str, model: str, api_key: str,
                         insecure_ssl: bool = False) -> Tuple[bool, str, float]:
    """Test a single provider with a real tiny request."""
    if not api_key:
        return False, "No API key provided", 0.0

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
    if insecure_ssl or os.environ.get("GATEWAY_INSECURE_SSL") == "1":
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=15, context=ctx) as resp:
            elapsed = (time.time() - t0) * 1000
            body = json.loads(resp.read().decode("utf-8", errors="replace"))
            reply = ""
            if body.get("choices"):
                msg = body["choices"][0].get("message") or {}
                reply = (msg.get("content") or msg.get("reasoning") or "").strip().replace("\n", " ")[:40]
            return True, reply or "OK", elapsed
    except urllib.error.HTTPError as e:
        elapsed = (time.time() - t0) * 1000
        try:
            err_detail = json.loads(e.read().decode("utf-8", errors="replace"))
            msg = err_detail.get("error", {}).get("message", e.reason)
        except Exception:
            msg = e.reason
        return False, f"HTTP {e.code}: {msg[:60]}", elapsed
    except Exception as e:
        elapsed = (time.time() - t0) * 1000
        return False, str(e)[:60], elapsed


# =============================================================================
# CLI IMPLEMENTATION
# =============================================================================

def print_banner():
    print(f"""
{C.BLUE}{C.BOLD}+----------------------------------------------------------+
|           ccgate - Claude Code Gateway  v{VERSION}           |
+----------------------------------------------------------+{C.RESET}
{C.DIM}Routes Claude Code through OpenAI-compatible providers{C.RESET}
""")


def cmd_doctor():
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
        if not api_key:
            print(f"  {C.YELLOW}[?]{C.RESET} {prov_name:<16} -- {C.DIM}No API key found in keys.xlsx, skipping{C.RESET}")
            continue

        ok, msg, elapsed = test_single_provider(prov_name, base_url, model, api_key, CONFIG.insecure_ssl)
        if ok:
            print(f"  {C.GREEN}[OK]{C.RESET} {prov_name:<16} {C.DIM}({model}){C.RESET} -- {C.GREEN}{elapsed:.0f}ms{C.RESET} \"{msg}\"")
        else:
            print(f"  {C.RED}[FAIL]{C.RESET} {prov_name:<16} {C.DIM}({model}){C.RESET} -- {C.RED}{msg}{C.RESET} ({elapsed:.0f}ms)")
    print()


def run_cli_interactive():
    print_banner()
    providers = load_providers_json()
    saved_keys = load_keys_from_excel()

    if not providers:
        print(f"{C.RED}No providers found in providers.json{C.RESET}")
        sys.exit(1)

    print(f"{C.BOLD}Select backend provider:{C.RESET}")
    pnames = list(providers.keys())
    for i, p in enumerate(pnames, 1):
        print(f"  [{i}] {p}")
    try:
        pidx = int(input(f"\nChoice [1-{len(pnames)}]: ")) - 1
        prov_name = pnames[pidx]
    except Exception:
        print(f"{C.RED}Invalid selection.{C.RESET}")
        sys.exit(1)

    prov_info = providers[prov_name]
    CONFIG.provider_name = prov_name
    CONFIG.base_url = prov_info.get("base_url", "")

    models = prov_info.get("models", [])
    if models:
        print(f"\n{C.BOLD}Select model:{C.RESET}")
        for i, m in enumerate(models, 1):
            print(f"  [{i}] {m}")
        print(f"  [{len(models) + 1}] Enter manually")
        try:
            midx = int(input(f"\nChoice [1-{len(models) + 1}]: ")) - 1
            if midx < len(models):
                CONFIG.backend_model = models[midx]
            else:
                CONFIG.backend_model = input("Enter model: ").strip()
        except Exception:
            CONFIG.backend_model = models[0]
    else:
        CONFIG.backend_model = input("Enter model name: ").strip()

    k = find_key_for_provider(prov_name, saved_keys)
    if k:
        print(f"\n{C.GREEN}[OK]{C.RESET} Using saved API key for {prov_name}")
        CONFIG.api_key = k
    else:
        CONFIG.api_key = input("Enter API key: ").strip()

    apply_claude_settings(CONFIG.port, CONFIG.host, CONFIG.backend_model)

    ok, msg = start_server_thread(CONFIG.port, CONFIG.host)
    if not ok:
        print(f"{C.RED}{msg}{C.RESET}")
        sys.exit(1)

    url = f"http://{CONFIG.host}:{CONFIG.port}"
    print(f"""
{C.GREEN}{C.BOLD}[OK] Gateway running{C.RESET} on {C.CYAN}{C.BOLD}{url}{C.RESET}

  Provider : {CONFIG.provider_name}
  Model    : {CONFIG.backend_model}
  Endpoint : {build_chat_url(CONFIG.base_url)}

Claude Code is configured! Run: {C.BOLD}claude{C.RESET} in your terminal.
Press Ctrl+C to stop.
{"-" * 60}""")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print(f"\n{C.YELLOW}Stopping gateway...{C.RESET}")
        stop_server_thread()
        print(f"{C.GREEN}Gateway stopped.{C.RESET}\n")


# =============================================================================
# PYSIDE6 GUI IMPLEMENTATION
# =============================================================================

def run_gui():
    """Launch the modern PySide6 GUI control panel."""
    try:
        from PySide6.QtWidgets import (
            QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
            QGridLayout, QLabel, QPushButton, QComboBox, QLineEdit, QSpinBox,
            QTextEdit, QFrame, QMessageBox, QListView
        )
        from PySide6.QtCore import Qt, QTimer, Signal, QObject
        from PySide6.QtGui import QFont, QIcon, QTextCursor
    except ImportError:
        print(f"{C.YELLOW}PySide6 is not installed. Falling back to interactive CLI.{C.RESET}")
        run_cli_interactive()
        return

    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")

    # Bridge for thread-safe UI logging
    class LogEmitter(QObject):
        log_signal = Signal(str, str, str, str, int, float, int, int)

    emitter = LogEmitter()

    def log_bridge(ts, method, path, model, status, elapsed, in_tok, out_tok):
        emitter.log_signal.emit(ts, method, path, model, status, elapsed, in_tok, out_tok)

    CONFIG.log_callback = log_bridge

    # Professional White Theme Stylesheet with Sapphire Blue, Emerald Green & Ruby Red Accents
    STYLESHEET = """
        QMainWindow {
            background-color: #f8fafc;
            color: #0f172a;
        }
        QWidget {
            color: #0f172a;
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            font-size: 13px;
        }
        QFrame#card {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
        }
        QLabel {
            color: #334155;
            font-weight: 500;
        }
        QLabel#title {
            color: #1e3a8a;
            font-size: 20px;
            font-weight: 800;
        }
        QLabel#brand_badge {
            background-color: #eff6ff;
            color: #2563eb;
            border: 1px solid #bfdbfe;
            border-radius: 5px;
            padding: 2px 7px;
            font-size: 11px;
            font-weight: 700;
        }
        QLabel#subtitle {
            color: #64748b;
            font-size: 12px;
        }
        QLabel#status_running {
            background-color: #ecfdf5;
            color: #047857;
            border: 1px solid #a7f3d0;
            border-radius: 13px;
            padding: 4px 12px;
            font-weight: 700;
            font-size: 12px;
        }
        QLabel#status_stopped {
            background-color: #fef2f2;
            color: #b91c1c;
            border: 1px solid #fecaca;
            border-radius: 13px;
            padding: 4px 12px;
            font-weight: 700;
            font-size: 12px;
        }
        QLabel#url_badge {
            background-color: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #bfdbfe;
            border-radius: 6px;
            padding: 4px 10px;
            font-family: 'Consolas', monospace;
            font-size: 12px;
            font-weight: 600;
        }
        QComboBox, QLineEdit, QSpinBox {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 7px 12px;
            color: #0f172a;
            selection-background-color: #2563eb;
            selection-color: #ffffff;
        }
        QComboBox:focus, QLineEdit:focus, QSpinBox:focus {
            border: 1.5px solid #2563eb;
            background-color: #ffffff;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            border-left: 1px solid #e2e8f0;
            border-top-right-radius: 6px;
            border-bottom-right-radius: 6px;
            background-color: #f8fafc;
        }
        QComboBox::down-arrow {
            width: 0;
            height: 0;
            border-left: 5px solid transparent;
            border-right: 5px solid transparent;
            border-top: 6px solid #64748b;
        }
        QComboBox::down-arrow:hover {
            border-top-color: #2563eb;
        }
        QComboBox QAbstractItemView, QListView {
            background-color: #ffffff;
            color: #0f172a;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 4px;
            selection-background-color: #2563eb;
            selection-color: #ffffff;
            outline: none;
        }
        QComboBox QAbstractItemView::item, QListView::item {
            min-height: 28px;
            padding: 5px 10px;
            border-radius: 4px;
            color: #0f172a;
        }
        QComboBox QAbstractItemView::item:hover, QListView::item:hover {
            background-color: #eff6ff;
            color: #1e40af;
        }
        QComboBox QAbstractItemView::item:selected, QListView::item:selected {
            background-color: #2563eb;
            color: #ffffff;
        }
        QPushButton {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: 600;
            color: #334155;
        }
        QPushButton:hover {
            background-color: #f8fafc;
            border-color: #94a3b8;
            color: #0f172a;
        }
        QPushButton:pressed {
            background-color: #f1f5f9;
        }
        /* Green Start Button */
        QPushButton#btn_start {
            background-color: #059669;
            border: 1px solid #047857;
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            padding: 9px 20px;
        }
        QPushButton#btn_start:hover {
            background-color: #10b981;
            border-color: #059669;
        }
        /* Red Stop Button */
        QPushButton#btn_stop {
            background-color: #dc2626;
            border: 1px solid #b91c1c;
            color: #ffffff;
            font-size: 13px;
            font-weight: 700;
            padding: 9px 20px;
        }
        QPushButton#btn_stop:hover {
            background-color: #ef4444;
            border-color: #dc2626;
        }
        /* Blue Claude Setup Button */
        QPushButton#btn_apply {
            background-color: #2563eb;
            border: 1px solid #1d4ed8;
            color: #ffffff;
            font-weight: 700;
            padding: 9px 18px;
        }
        QPushButton#btn_apply:hover {
            background-color: #1d4ed8;
        }
        /* Test Button with Indigo / Blue text */
        QPushButton#btn_test {
            background-color: #f8fafc;
            border: 1px solid #cbd5e1;
            color: #1e40af;
            font-weight: 600;
            padding: 9px 16px;
        }
        QPushButton#btn_test:hover {
            background-color: #eff6ff;
            border-color: #93c5fd;
            color: #1d4ed8;
        }
        /* Log Area - Light Crisp Developer Console */
        QTextEdit {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 12px;
            color: #0f172a;
            padding: 10px;
        }
    """

    class CcgateWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle(f"ccgate — Claude Code Gateway v{VERSION}")
            self.setMinimumSize(780, 720)
            self.setStyleSheet(STYLESHEET)

            self.providers = load_providers_json()
            self.saved_keys = load_keys_from_excel()

            self.init_ui()
            self.populate_data()

            # Connect log signal
            emitter.log_signal.connect(self.append_log)

            # Auto-update status timer
            self.timer = QTimer(self)
            self.timer.timeout.connect(self.update_status_display)
            self.timer.start(1000)

        def init_ui(self):
            central = QWidget()
            self.setCentralWidget(central)
            main_layout = QVBoxLayout(central)
            main_layout.setContentsMargins(18, 18, 18, 18)
            main_layout.setSpacing(14)

            # --- 1. Header Card ---
            header_card = QFrame()
            header_card.setObjectName("card")
            h_layout = QHBoxLayout(header_card)
            h_layout.setContentsMargins(16, 12, 16, 12)

            title_box = QVBoxLayout()
            title_row = QHBoxLayout()
            title_lbl = QLabel("ccgate")
            title_lbl.setObjectName("title")
            badge_lbl = QLabel(f"v{VERSION}")
            badge_lbl.setObjectName("brand_badge")
            title_row.addWidget(title_lbl)
            title_row.addWidget(badge_lbl)
            title_row.addStretch()

            sub_lbl = QLabel("Self-hosted reverse proxy for Claude Code (Anthropic <-> OpenAI)")
            sub_lbl.setObjectName("subtitle")
            title_box.addLayout(title_row)
            title_box.addWidget(sub_lbl)
            h_layout.addLayout(title_box)

            h_layout.addStretch()

            status_box = QVBoxLayout()
            status_box.setAlignment(Qt.AlignRight)
            self.status_lbl = QLabel("● Stopped")
            self.status_lbl.setObjectName("status_stopped")
            self.info_lbl = QLabel("Port: 5005 | Req: 0")
            self.info_lbl.setObjectName("subtitle")
            status_box.addWidget(self.status_lbl)
            status_box.addWidget(self.info_lbl)
            h_layout.addLayout(status_box)

            main_layout.addWidget(header_card)

            # --- 2. Configuration Card ---
            config_card = QFrame()
            config_card.setObjectName("card")
            c_layout = QGridLayout(config_card)
            c_layout.setContentsMargins(16, 16, 16, 16)
            c_layout.setHorizontalSpacing(14)
            c_layout.setVerticalSpacing(12)

            # Provider
            c_layout.addWidget(QLabel("Backend Provider:"), 0, 0)
            self.prov_combo = QComboBox()
            self.prov_combo.setView(QListView())
            self.prov_combo.setMaxVisibleItems(12)
            self.prov_combo.currentIndexChanged.connect(self.on_provider_changed)
            c_layout.addWidget(self.prov_combo, 0, 1)

            # Port
            c_layout.addWidget(QLabel("Gateway Port:"), 0, 2)
            self.port_spin = QSpinBox()
            self.port_spin.setRange(1024, 65535)
            self.port_spin.setValue(DEFAULT_PORT)
            c_layout.addWidget(self.port_spin, 0, 3)

            # Model
            c_layout.addWidget(QLabel("Backend Model:"), 1, 0)
            self.model_combo = QComboBox()
            self.model_combo.setEditable(True)
            self.model_combo.setView(QListView())
            self.model_combo.setMaxVisibleItems(12)
            c_layout.addWidget(self.model_combo, 1, 1, 1, 3)

            # API Key
            c_layout.addWidget(QLabel("API Key:"), 2, 0)
            key_layout = QHBoxLayout()
            self.key_edit = QLineEdit()
            self.key_edit.setEchoMode(QLineEdit.Password)
            self.key_edit.setPlaceholderText("Paste or select key from keys.xlsx...")
            key_layout.addWidget(self.key_edit)

            self.btn_toggle_key = QPushButton("👁")
            self.btn_toggle_key.setFixedWidth(38)
            self.btn_toggle_key.clicked.connect(self.toggle_key_visibility)
            key_layout.addWidget(self.btn_toggle_key)

            c_layout.addLayout(key_layout, 2, 1, 1, 3)

            # Base URL info
            c_layout.addWidget(QLabel("Endpoint URL:"), 3, 0)
            self.url_lbl = QLabel("")
            self.url_lbl.setObjectName("url_badge")
            c_layout.addWidget(self.url_lbl, 3, 1, 1, 3)

            main_layout.addWidget(config_card)

            # --- 3. Actions Row ---
            act_layout = QHBoxLayout()
            act_layout.setSpacing(10)

            self.btn_toggle_server = QPushButton("▶ Start Gateway")
            self.btn_toggle_server.setObjectName("btn_start")
            self.btn_toggle_server.clicked.connect(self.toggle_server)
            act_layout.addWidget(self.btn_toggle_server)

            self.btn_apply_claude = QPushButton("⚡ Configure Claude Code")
            self.btn_apply_claude.setObjectName("btn_apply")
            self.btn_apply_claude.setToolTip("Sets ANTHROPIC_BASE_URL to http://127.0.0.1:5005 in ~/.claude/settings.json")
            self.btn_apply_claude.clicked.connect(self.on_apply_claude)
            act_layout.addWidget(self.btn_apply_claude)

            self.btn_test = QPushButton("🩺 Test Provider")
            self.btn_test.setObjectName("btn_test")
            self.btn_test.clicked.connect(self.on_test_provider)
            act_layout.addWidget(self.btn_test)

            self.btn_clear_log = QPushButton("Clear Log")
            self.btn_clear_log.clicked.connect(lambda: self.log_text.clear())
            act_layout.addWidget(self.btn_clear_log)

            main_layout.addLayout(act_layout)

            # --- 4. Live Traffic Log Card ---
            log_card = QFrame()
            log_card.setObjectName("card")
            l_layout = QVBoxLayout(log_card)
            l_layout.setContentsMargins(14, 14, 14, 14)

            log_header = QHBoxLayout()
            log_title = QLabel("Live Traffic Monitor")
            log_title.setStyleSheet("font-weight: 700; color: #1e293b; font-size: 13px;")
            log_header.addWidget(log_title)
            log_header.addStretch()

            self.copy_btn = QPushButton("📋 Copy URL")
            self.copy_btn.setStyleSheet("color: #2563eb; font-weight: 600; padding: 4px 10px;")
            self.copy_btn.clicked.connect(self.copy_url)
            log_header.addWidget(self.copy_btn)

            l_layout.addLayout(log_header)

            self.log_text = QTextEdit()
            self.log_text.setReadOnly(True)
            l_layout.addWidget(self.log_text)

            main_layout.addWidget(log_card)

        def populate_data(self):
            """Load providers and saved keys into the UI."""
            self.prov_combo.clear()
            for p in self.providers.keys():
                self.prov_combo.addItem(p)

            # Default to Openrouter or Groq if present
            for target in ["Openrouter", "groq", "NVIDIA"]:
                idx = self.prov_combo.findText(target, Qt.MatchFixedString)
                if idx >= 0:
                    self.prov_combo.setCurrentIndex(idx)
                    break

            self.on_provider_changed()

        def on_provider_changed(self):
            """Update model list and key when provider changes."""
            prov_name = self.prov_combo.currentText().strip()
            prov_info = self.providers.get(prov_name, {})
            base_url = prov_info.get("base_url", "")
            self.url_lbl.setText(build_chat_url(base_url))

            self.model_combo.clear()
            models = prov_info.get("models", [])
            for m in models:
                self.model_combo.addItem(m)
            if models:
                self.model_combo.setCurrentIndex(0)

            # Auto-find API key
            k = find_key_for_provider(prov_name, self.saved_keys)
            if k:
                self.key_edit.setText(k)

        def toggle_key_visibility(self):
            if self.key_edit.echoMode() == QLineEdit.Password:
                self.key_edit.setEchoMode(QLineEdit.Normal)
                self.btn_toggle_key.setText("🔒")
            else:
                self.key_edit.setEchoMode(QLineEdit.Password)
                self.btn_toggle_key.setText("👁")

        def copy_url(self):
            url = f"http://{DEFAULT_HOST}:{self.port_spin.value()}"
            QApplication.clipboard().setText(url)
            self.log_text.append(f"<span style='color:#2563eb; font-weight:600;'>[INFO] Copied {url} to clipboard</span>")

        def toggle_server(self):
            """Start or stop the gateway."""
            if CONFIG.is_running:
                ok, msg = stop_server_thread()
                self.status_lbl.setText("● Stopped")
                self.status_lbl.setObjectName("status_stopped")
                self.status_lbl.setStyle(self.status_lbl.style())
                self.btn_toggle_server.setText("▶ Start Gateway")
                self.btn_toggle_server.setObjectName("btn_start")
                self.btn_toggle_server.setStyle(self.btn_toggle_server.style())
                self.log_text.append(f"<span style='color:#d97706; font-weight:600;'>[GATEWAY] {msg}</span>")
            else:
                prov_name = self.prov_combo.currentText().strip()
                prov_info = self.providers.get(prov_name, {})
                CONFIG.provider_name = prov_name
                CONFIG.base_url = prov_info.get("base_url", "")
                CONFIG.backend_model = self.model_combo.currentText().strip()
                CONFIG.api_key = self.key_edit.text().strip()
                port = self.port_spin.value()

                if not CONFIG.backend_model:
                    QMessageBox.warning(self, "Missing Model", "Please specify or select a backend model.")
                    return
                if not CONFIG.api_key:
                    QMessageBox.warning(self, "Missing Key", "Please provide an API key.")
                    return

                ok, msg = start_server_thread(port, DEFAULT_HOST)
                if ok:
                    url = f"http://{DEFAULT_HOST}:{port}"
                    self.status_lbl.setText(f"● Running on {url}")
                    self.status_lbl.setObjectName("status_running")
                    self.status_lbl.setStyle(self.status_lbl.style())
                    self.btn_toggle_server.setText("■ Stop Gateway")
                    self.btn_toggle_server.setObjectName("btn_stop")
                    self.btn_toggle_server.setStyle(self.btn_toggle_server.style())
                    self.log_text.append(f"<span style='color:#059669; font-weight:bold;'>[GATEWAY] {msg}</span>")
                    self.log_text.append(f"<span style='color:#475569;'>  Model: <b style='color:#0f172a;'>{CONFIG.backend_model}</b> | Provider: <b style='color:#0f172a;'>{CONFIG.provider_name}</b></span>")
                else:
                    QMessageBox.critical(self, "Startup Error", msg)
                    self.log_text.append(f"<span style='color:#dc2626; font-weight:bold;'>[ERROR] {msg}</span>")

        def on_apply_claude(self):
            """Configure ~/.claude/settings.json."""
            port = self.port_spin.value()
            model = self.model_combo.currentText().strip()
            ok, msg = apply_claude_settings(port, DEFAULT_HOST, model)
            if ok:
                QMessageBox.information(self, "Claude Code Configured",
                                        f"{msg}\n\nYou can now run 'claude' in your terminal!")
                self.log_text.append(f"<span style='color:#2563eb; font-weight:bold;'>[CONFIG] {msg}</span>")
            else:
                QMessageBox.critical(self, "Config Error", msg)

        def on_test_provider(self):
            """Test the selected provider connectivity."""
            prov_name = self.prov_combo.currentText().strip()
            prov_info = self.providers.get(prov_name, {})
            base_url = prov_info.get("base_url", "")
            model = self.model_combo.currentText().strip()
            api_key = self.key_edit.text().strip()

            if not api_key:
                QMessageBox.warning(self, "Missing Key", "Please enter an API key to test.")
                return

            self.log_text.append(f"<span style='color:#64748b;'>[TEST] Testing {prov_name} ({model})...</span>")
            self.btn_test.setEnabled(False)
            QApplication.processEvents()

            ok, msg, elapsed = test_single_provider(prov_name, base_url, model, api_key, CONFIG.insecure_ssl)
            self.btn_test.setEnabled(True)

            if ok:
                self.log_text.append(f"<span style='color:#059669; font-weight:bold;'>[TEST OK] {prov_name} ({elapsed:.0f}ms): \"{msg}\"</span>")
            else:
                self.log_text.append(f"<span style='color:#dc2626; font-weight:bold;'>[TEST FAIL] {prov_name} ({elapsed:.0f}ms): {msg}</span>")

        def append_log(self, ts, method, path, model, status, elapsed, in_tok, out_tok):
            status_color = "#059669" if 200 <= status < 300 else "#dc2626"
            method_color = "#2563eb"
            line = (
                f"<span style='color:#64748b; font-weight:600;'>{ts}</span> "
                f"<span style='color:{method_color}; font-weight:700;'>{method}</span> "
                f"<span style='color:#0f172a;'>{path}</span> "
                f"<span style='color:#7c3aed; font-weight:600;'>-> {model}</span> "
                f"<span style='color:{status_color}; font-weight:bold;'>[{status}]</span> "
                f"<span style='color:#64748b;'>({elapsed:.0f}ms, {in_tok}in/{out_tok}out)</span>"
            )
            self.log_text.append(line)
            self.log_text.moveCursor(QTextCursor.End)

        def update_status_display(self):
            if CONFIG.is_running:
                uptime = int(time.time() - CONFIG.start_time)
                self.info_lbl.setText(f"Up: {uptime}s | Req: {CONFIG.request_count}")
            else:
                self.info_lbl.setText(f"Port: {self.port_spin.value()} | Req: 0")

        def closeEvent(self, event):
            """Ensure background server thread is cleanly shut down when closing window."""
            stop_server_thread()
            event.accept()

    win = CcgateWindow()
    win.show()
    sys.exit(app.exec())


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="ccgate - Claude Code Gateway",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("command", nargs="?", default=None, help="Command: setup, doctor")
    parser.add_argument("--cli", action="store_true", help="Run interactive terminal CLI instead of GUI")
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

    # 1. Standalone commands
    if args.command == "doctor":
        cmd_doctor()
        return

    if args.command == "setup":
        print_banner()
        apply_claude_settings(CONFIG.port, CONFIG.host, args.model)
        return

    # 2. CLI direct start mode with --provider
    if args.provider:
        print_banner()
        providers = load_providers_json()
        saved_keys = load_keys_from_excel()

        matched = None
        for pname in providers:
            if pname.strip().lower() == args.provider.strip().lower():
                matched = pname
                break
        if not matched:
            print(f"{C.RED}Provider '{args.provider}' not found in providers.json{C.RESET}")
            sys.exit(1)

        CONFIG.provider_name = matched
        CONFIG.base_url = args.base_url or providers[matched].get("base_url", "")
        models = providers[matched].get("models", [])
        CONFIG.backend_model = args.model or (models[0] if models else "")

        if args.key:
            CONFIG.api_key = args.key.strip()
        else:
            CONFIG.api_key = find_key_for_provider(matched, saved_keys)

        if not CONFIG.api_key:
            print(f"{C.RED}No API key found for {matched}.{C.RESET}")
            sys.exit(1)

        if args.setup:
            apply_claude_settings(CONFIG.port, CONFIG.host, CONFIG.backend_model)

        ok, msg = start_server_thread(CONFIG.port, CONFIG.host)
        if not ok:
            print(f"{C.RED}{msg}{C.RESET}")
            sys.exit(1)

        url = f"http://{CONFIG.host}:{CONFIG.port}"
        print(f"""
{C.GREEN}{C.BOLD}[OK] Gateway running{C.RESET} on {C.CYAN}{C.BOLD}{url}{C.RESET}

  Provider : {CONFIG.provider_name}
  Model    : {CONFIG.backend_model}
  Endpoint : {build_chat_url(CONFIG.base_url)}

Claude Code is configured! Run: {C.BOLD}claude{C.RESET} in your terminal.
Press Ctrl+C to stop.
{"-" * 60}""")

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            stop_server_thread()
            print(f"{C.GREEN}Gateway stopped.{C.RESET}\n")
        return

    # 3. Interactive CLI mode if explicitly requested
    if args.cli:
        run_cli_interactive()
        return

    # 4. Default: Launch rich PySide6 GUI
    run_gui()


if __name__ == "__main__":
    main()

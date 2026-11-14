"""
===============================================================================
Key-Switcher: GUI API Tester
===============================================================================
A lightweight GUI tool to verify Base URL, API Key, and Model connectivity
for LLM providers (OpenAI-compatible, Anthropic-compatible, Ollama, OpenRouter,
Hugging Face, NVIDIA, etc.).

Features:
- Test any Base URL, API Key, and Model with 1 click.
- Auto-detects OpenAI vs Anthropic endpoint formats & normalizes URLs.
- Clear visual status: WORKING (Green) vs NOT WORKING (Red) with response time.
- Displays model response text, token usage, latency, and full raw JSON.
- Integrates with providers.json and keys.xlsx for quick-fill presets.
- Non-blocking asynchronous network requests with cancel support.
- Primary GUI: PySide6 (matches Key-Switcher).
- Fallback GUI: Tkinter (runs out of the box with standard Python library).
===============================================================================
"""

import sys
import os
import json
import time
import ssl
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional, Tuple, List


# =============================================================================
# FILE PATHS & CONSTANTS
# =============================================================================

PROJECT_DIR = Path(__file__).resolve().parent
PROVIDERS_FILE = PROJECT_DIR / "providers.json"
KEY_FILE = Path(r"C:\My Space\Data\keys.xlsx")

DEFAULT_PROMPT = "Respond with: API is working properly!"
DEFAULT_MAX_TOKENS = 32
DEFAULT_TIMEOUT = 25


# =============================================================================
# CORE API TEST ENGINE (Pure Python Standard Library - Zero Dependencies)
# =============================================================================

class ApiTestResult:
    """Holds structured test results from an API verification attempt."""

    def __init__(
        self,
        is_working: bool,
        status_code: int,
        status_message: str,
        latency_ms: float,
        endpoint_used: str,
        protocol: str,
        model_reply: str = "",
        token_usage: str = "",
        raw_response: str = "",
        response_headers: Optional[Dict[str, str]] = None,
        error_details: str = "",
        troubleshooting: str = "",
    ):
        self.is_working = is_working
        self.status_code = status_code
        self.status_message = status_message
        self.latency_ms = latency_ms
        self.endpoint_used = endpoint_used
        self.protocol = protocol
        self.model_reply = model_reply
        self.token_usage = token_usage
        self.raw_response = raw_response
        self.response_headers = response_headers or {}
        self.error_details = error_details
        self.troubleshooting = troubleshooting
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def to_markdown_report(self, base_url: str, model: str) -> str:
        status_badge = "✅ WORKING" if self.is_working else "❌ NOT WORKING"
        lines = [
            f"# API Test Report - {status_badge}",
            f"- **Timestamp**: {self.timestamp}",
            f"- **Status Code**: {self.status_code} ({self.status_message})",
            f"- **Latency**: {self.latency_ms:.1f} ms",
            f"- **Base URL**: `{base_url}`",
            f"- **Endpoint Tested**: `{self.endpoint_used}`",
            f"- **Model**: `{model}`",
            f"- **Protocol**: {self.protocol}",
        ]
        if self.token_usage:
            lines.append(f"- **Token Usage**: {self.token_usage}")
        if self.model_reply:
            lines.append(f"\n### Model Response\n> {self.model_reply}")
        if self.error_details:
            lines.append(f"\n### Error Details\n```text\n{self.error_details}\n```")
        if self.troubleshooting:
            lines.append(f"\n### Troubleshooting Advice\n💡 {self.troubleshooting}")
        if self.raw_response:
            lines.append(f"\n### Raw Response Body\n```json\n{self.raw_response[:3000]}\n```")
        return "\n".join(lines)


def get_troubleshooting_advice(status_code: int, error_text: str, endpoint: str) -> str:
    """Provides actionable advice based on HTTP status code and error messages."""
    error_lower = error_text.lower()
    
    if status_code == 401:
        return (
            "Authentication Failed (401): The API Key is invalid, expired, or missing "
            "permissions for this provider. Verify that the API Key was copied correctly."
        )
    elif status_code == 403:
        return (
            "Access Forbidden (403): Your account or API key does not have access "
            "to this model/endpoint, or geographic restrictions/IP block apply."
        )
    elif status_code == 404:
        return (
            "Not Found (404): Either the model name is incorrect, or the endpoint URL is wrong. "
            "Verify that the model exists on this provider and check the Base URL."
        )
    elif status_code == 429:
        return (
            "Rate Limited or Insufficient Credits (429): You have exceeded your rate limit, "
            "or your account has run out of credits/tokens. Check provider balance."
        )
    elif status_code == 400:
        return (
            "Bad Request (400): The model might not support the requested parameters, "
            "or the model identifier is malformed."
        )
    elif status_code in (500, 502, 503, 504):
        return (
            f"Provider Server Error ({status_code}): The upstream provider server is "
            "experiencing temporary downtime, high load, or a gateway error."
        )
    elif "certificate_verify_failed" in error_lower:
        return (
            "SSL Certificate Verification Failed: Your network or proxy intercepted the request. "
            "Try disabling 'Verify SSL' in Advanced Settings if using a self-signed proxy."
        )
    elif "timed out" in error_lower:
        return (
            "Connection Timed Out: The server did not respond in time. Check the host or "
            "increase the timeout in Advanced Settings."
        )
    elif "getaddrinfo failed" in error_lower or "name resolution" in error_lower:
        return "DNS Lookup Failed: Hostname could not be resolved. Check Base URL spelling."
    elif "connection refused" in error_lower:
        return "Connection Refused: Target port is closed or server is not running (e.g. Ollama offline)."
    
    return "Check Base URL, API Key, and Model name for typographical errors."


def extract_model_reply(data: Any) -> str:
    """Extracts the generated message text from OpenAI, Anthropic, or Ollama response formats."""
    if not isinstance(data, dict):
        return str(data)

    # 1. OpenAI-compatible format (choices[0].message.content)
    choices = data.get("choices")
    if isinstance(choices, list) and choices:
        choice = choices[0]
        if isinstance(choice, dict):
            msg = choice.get("message")
            if isinstance(msg, dict):
                content = msg.get("content")
                if isinstance(content, str):
                    return content.strip()
                elif isinstance(content, list):
                    # Multi-part content
                    parts = []
                    for part in content:
                        if isinstance(part, dict) and "text" in part:
                            parts.append(part["text"])
                    if parts:
                        return "".join(parts).strip()
            # Legacy completions format (choices[0].text)
            text = choice.get("text")
            if isinstance(text, str):
                return text.strip()

    # 2. Anthropic-compatible format (content[0].text)
    content = data.get("content")
    if isinstance(content, list) and content:
        texts = [p.get("text", "") for p in content if isinstance(p, dict) and p.get("type") == "text"]
        if texts:
            return "".join(texts).strip()
    elif isinstance(content, str):
        return content.strip()

    # 3. Ollama format (response or message.content)
    if "response" in data and isinstance(data["response"], str):
        return data["response"].strip()
    if "message" in data and isinstance(data["message"], dict):
        return str(data["message"].get("content", "")).strip()

    # Fallback: check top-level text/generated_text
    if "text" in data and isinstance(data["text"], str):
        return data["text"].strip()
    if "generated_text" in data and isinstance(data["generated_text"], str):
        return data["generated_text"].strip()

    return ""


def extract_token_usage(data: Any) -> str:
    """Extracts token usage metrics from the response."""
    if not isinstance(data, dict):
        return ""
    usage = data.get("usage")
    if isinstance(usage, dict):
        prompt_tokens = usage.get("prompt_tokens") or usage.get("input_tokens")
        completion_tokens = usage.get("completion_tokens") or usage.get("output_tokens")
        total_tokens = usage.get("total_tokens")
        parts = []
        if prompt_tokens is not None:
            parts.append(f"Prompt: {prompt_tokens}")
        if completion_tokens is not None:
            parts.append(f"Completion: {completion_tokens}")
        if total_tokens is not None:
            parts.append(f"Total: {total_tokens}")
        if parts:
            return " | ".join(parts)
    return ""


def extract_error_message(data: Any) -> str:
    """Extracts detailed error message from response JSON."""
    if isinstance(data, dict):
        err = data.get("error")
        if isinstance(err, dict):
            msg = err.get("message") or err.get("detail") or err.get("code")
            if msg:
                return str(msg)
            return json.dumps(err)
        elif isinstance(err, str):
            return err
        for key in ("detail", "message", "msg", "error_description"):
            if key in data:
                return str(data[key])
    return ""


def build_candidate_endpoints(base_url: str, protocol: str) -> List[Tuple[str, str]]:
    """
    Returns a list of (url, protocol_style) tuples to test.
    protocol_style: 'openai' or 'anthropic'
    """
    clean_url = base_url.strip().rstrip("/")
    if not clean_url:
        return []

    # Auto-add scheme if missing (e.g. user typed openrouter.ai/api or localhost:11434)
    if not clean_url.startswith("http://") and not clean_url.startswith("https://"):
        if clean_url.startswith("localhost") or clean_url.startswith("127.0.0.1"):
            clean_url = "http://" + clean_url
        else:
            clean_url = "https://" + clean_url

    # If protocol is forced to Direct URL, use as-is with openai payload style
    if protocol == "direct":
        return [(clean_url, "openai")]

    # If user selected Anthropic explicitly
    if protocol == "anthropic":
        if clean_url.endswith("/messages"):
            return [(clean_url, "anthropic")]
        if clean_url.endswith("/v1"):
            return [(f"{clean_url}/messages", "anthropic")]
        return [
            (f"{clean_url}/v1/messages", "anthropic"),
            (f"{clean_url}/messages", "anthropic"),
        ]

    # If user selected OpenAI explicitly
    if protocol == "openai":
        if clean_url.endswith("/chat/completions"):
            return [(clean_url, "openai")]
        if clean_url.endswith("/v1"):
            return [(f"{clean_url}/chat/completions", "openai")]
        return [
            (f"{clean_url}/v1/chat/completions", "openai"),
            (f"{clean_url}/chat/completions", "openai"),
        ]

    # AUTO-DETECT MODE:
    # 1. If URL contains /messages or anthropic
    if clean_url.endswith("/messages") or "anthropic" in clean_url.lower():
        if clean_url.endswith("/messages"):
            return [(clean_url, "anthropic")]
        if clean_url.endswith("/v1"):
            return [(f"{clean_url}/messages", "anthropic")]
        return [
            (f"{clean_url}/v1/messages", "anthropic"),
            (f"{clean_url}/messages", "anthropic"),
        ]

    # 2. If URL already ends with /chat/completions
    if clean_url.endswith("/chat/completions"):
        return [(clean_url, "openai")]

    # 3. Standard OpenAI candidates first, then Anthropic candidates
    candidates = []
    if clean_url.endswith("/v1"):
        candidates.append((f"{clean_url}/chat/completions", "openai"))
        candidates.append((f"{clean_url}/messages", "anthropic"))
    elif clean_url.endswith("/api"):
        candidates.append((f"{clean_url}/v1/chat/completions", "openai"))
        candidates.append((f"{clean_url}/chat/completions", "openai"))
    else:
        candidates.append((f"{clean_url}/v1/chat/completions", "openai"))
        candidates.append((f"{clean_url}/chat/completions", "openai"))
        candidates.append((f"{clean_url}/v1/messages", "anthropic"))
        candidates.append((f"{clean_url}/messages", "anthropic"))

    return candidates


def perform_single_api_call(
    endpoint_url: str,
    protocol_style: str,
    api_key: str,
    model: str,
    prompt: str,
    max_tokens: int,
    timeout: int,
    verify_ssl: bool,
) -> Tuple[bool, int, str, float, str, Dict[str, str]]:
    """
    Executes a single HTTP POST request.
    Returns: (is_success, status_code, response_body_text, latency_ms, status_phrase, headers)
    """
    # 1. Build Payload
    if protocol_style == "anthropic":
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_tokens,
        }
    else:
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": max_tokens,
            "temperature": 0.2,
        }

    post_data = json.dumps(payload).encode("utf-8")

    # 2. Build Headers
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "KeySwitcher-APITester/1.0",
    }
    if api_key.strip():
        # OpenRouter / OpenAI standard
        headers["Authorization"] = f"Bearer {api_key.strip()}"
        # Anthropic standard
        if protocol_style == "anthropic":
            headers["x-api-key"] = api_key.strip()
            headers["anthropic-version"] = "2023-06-01"
        # OpenRouter metadata (optional, best practice)
        if "openrouter.ai" in endpoint_url:
            headers["HTTP-Referer"] = "https://github.com/Key-Switcher"
            headers["X-Title"] = "Key Switcher API Tester"

    req = urllib.request.Request(endpoint_url, data=post_data, headers=headers, method="POST")

    # 3. SSL Context
    ssl_context = ssl.create_default_context()
    if not verify_ssl:
        ssl_context.check_hostname = False
        ssl_context.verify_mode = ssl.CERT_NONE

    start_time = time.perf_counter()
    status_code = 0
    status_phrase = "Unknown"
    body_text = ""
    resp_headers: Dict[str, str] = {}

    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_context) as resp:
            latency_ms = (time.perf_counter() - start_time) * 1000.0
            status_code = resp.status
            status_phrase = resp.reason or "OK"
            body_text = resp.read().decode("utf-8", errors="replace")
            for k, v in resp.headers.items():
                resp_headers[k] = v
            return (True, status_code, body_text, latency_ms, status_phrase, resp_headers)

    except urllib.error.HTTPError as err:
        latency_ms = (time.perf_counter() - start_time) * 1000.0
        status_code = err.code
        status_phrase = err.reason or "HTTP Error"
        try:
            body_text = err.read().decode("utf-8", errors="replace")
        except Exception:
            body_text = str(err)
        if hasattr(err, "headers") and err.headers:
            for k, v in err.headers.items():
                resp_headers[k] = v
        return (False, status_code, body_text, latency_ms, status_phrase, resp_headers)

    except urllib.error.URLError as err:
        latency_ms = (time.perf_counter() - start_time) * 1000.0
        reason_str = str(err.reason)
        return (False, 0, f"Network/URL Error: {reason_str}", latency_ms, reason_str, {})

    except TimeoutError:
        latency_ms = (time.perf_counter() - start_time) * 1000.0
        return (False, 408, "Request Timed Out", latency_ms, "Request Timed Out", {})

    except Exception as exc:
        latency_ms = (time.perf_counter() - start_time) * 1000.0
        return (False, 0, f"Unexpected Error: {str(exc)}", latency_ms, str(exc), {})


def run_api_test(
    base_url: str,
    api_key: str,
    model: str,
    protocol: str = "auto",
    prompt: str = DEFAULT_PROMPT,
    max_tokens: int = DEFAULT_MAX_TOKENS,
    timeout: int = DEFAULT_TIMEOUT,
    verify_ssl: bool = True,
) -> ApiTestResult:
    """
    Main entry point for testing an API configuration.
    Tries candidate endpoints if auto-detecting.
    """
    base_url = base_url.strip()
    model = model.strip()

    if not base_url:
        return ApiTestResult(
            is_working=False,
            status_code=0,
            status_message="Missing Base URL",
            latency_ms=0,
            endpoint_used="",
            protocol=protocol,
            error_details="Base URL cannot be empty.",
            troubleshooting="Please enter a valid Base URL (e.g. https://openrouter.ai/api/v1).",
        )

    if not model:
        return ApiTestResult(
            is_working=False,
            status_code=0,
            status_message="Missing Model",
            latency_ms=0,
            endpoint_used="",
            protocol=protocol,
            error_details="Model name cannot be empty.",
            troubleshooting="Please enter or select a model name to test (e.g. openai/gpt-oss-120b).",
        )

    candidates = build_candidate_endpoints(base_url, protocol)
    if not candidates:
        return ApiTestResult(
            is_working=False,
            status_code=0,
            status_message="Invalid URL",
            latency_ms=0,
            endpoint_used=base_url,
            protocol=protocol,
            error_details=f"Could not build candidate endpoints from URL '{base_url}'.",
        )

    last_failure_result: Optional[ApiTestResult] = None

    for endpoint_url, protocol_style in candidates:
        is_success, status_code, body_text, latency_ms, status_phrase, headers = perform_single_api_call(
            endpoint_url=endpoint_url,
            protocol_style=protocol_style,
            api_key=api_key,
            model=model,
            prompt=prompt,
            max_tokens=max_tokens,
            timeout=timeout,
            verify_ssl=verify_ssl,
        )

        # Attempt to parse body as JSON
        parsed_json = None
        try:
            parsed_json = json.loads(body_text)
        except Exception:
            parsed_json = None

        if is_success and status_code == 200:
            # Model output and token usage
            model_reply = extract_model_reply(parsed_json) if parsed_json else body_text.strip()
            token_usage = extract_token_usage(parsed_json) if parsed_json else ""
            raw_formatted = json.dumps(parsed_json, indent=2) if parsed_json else body_text

            return ApiTestResult(
                is_working=True,
                status_code=200,
                status_message="OK - Working",
                latency_ms=latency_ms,
                endpoint_used=endpoint_url,
                protocol=f"{protocol_style.upper()} ({protocol})",
                model_reply=model_reply,
                token_usage=token_usage,
                raw_response=raw_formatted,
                response_headers=headers,
                troubleshooting="Everything is working! This configuration can be used in Key-Switcher.",
            )

        # Handle failure
        error_msg = extract_error_message(parsed_json) if parsed_json else body_text
        advice = get_troubleshooting_advice(status_code, error_msg or body_text, endpoint_url)
        raw_formatted = json.dumps(parsed_json, indent=2) if parsed_json else body_text

        result = ApiTestResult(
            is_working=False,
            status_code=status_code,
            status_message=status_phrase,
            latency_ms=latency_ms,
            endpoint_used=endpoint_url,
            protocol=f"{protocol_style.upper()} ({protocol})",
            error_details=error_msg or status_phrase,
            raw_response=raw_formatted,
            response_headers=headers,
            troubleshooting=advice,
        )

        last_failure_result = result

        # If 404 and we have more candidates in auto mode, keep trying next candidate
        if status_code == 404 and len(candidates) > 1:
            continue
        else:
            # For 401, 403, 429, etc., don't loop through further candidates; return immediate failure
            break

    return last_failure_result or ApiTestResult(
        is_working=False,
        status_code=0,
        status_message="Test Failed",
        latency_ms=0,
        endpoint_used=base_url,
        protocol=protocol,
        error_details="No endpoints could be verified.",
    )


# =============================================================================
# DATA HELPERS (providers.json & keys.xlsx)
# =============================================================================

def find_key_file() -> Path:
    candidates = [
        Path(r"C:\My Space\Data\keys.xlsx"),
        PROJECT_DIR / "keys.xlsx",
        Path.cwd() / "keys.xlsx",
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


def find_providers_file() -> Path:
    candidates = [
        PROJECT_DIR / "providers.json",
        Path.cwd() / "providers.json",
        Path(r"C:\My Space\GitHub Repo\Key-Switcher\providers.json"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


def load_providers_json() -> Dict[str, Any]:
    """Loads provider presets from providers.json."""
    pf = find_providers_file()
    if not pf.exists():
        return {}
    try:
        content = pf.read_text(encoding="utf-8")
        data = json.loads(content)
        return data.get("providers", {})
    except Exception:
        return {}


def load_keys_from_excel() -> Dict[str, Dict[str, str]]:
    """
    Loads saved API keys from keys.xlsx with robust dual-method reading:
    Method 1: openpyxl
    Method 2: Pure Python standard library (zipfile + xml.etree.ElementTree)
    Returns: { nickname: { 'provider': ..., 'key': ... } }
    """
    kf = find_key_file()
    if not kf.exists():
        return {}

    # Method 1: openpyxl
    try:
        import openpyxl
        wb = openpyxl.load_workbook(kf, read_only=True, data_only=True)
        ws = wb.active
        headers = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, ws.max_column + 1)]

        provider_col = headers.index("provider") + 1 if "provider" in headers else 1
        nickname_col = headers.index("nickname") + 1 if "nickname" in headers else 2
        key_col = headers.index("api key") + 1 if "api key" in headers else (
            headers.index("key") + 1 if "key" in headers else 3
        )

        keys_data = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < nickname_col:
                continue
            nick = str(row[nickname_col - 1] or "").strip()
            if not nick:
                continue
            prov = str(row[provider_col - 1] or "").strip() if len(row) >= provider_col else ""
            key = str(row[key_col - 1] or "").strip() if len(row) >= key_col else ""
            if key:
                if key.lower().startswith("bearer "):
                    key = key[7:].strip()
                keys_data[nick] = {"provider": prov, "key": key}
        wb.close()
        if keys_data:
            return keys_data
    except Exception:
        pass

    # Method 2: Pure Python standard library parser (Zero dependencies)
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(kf, "r") as z:
            # Shared strings
            shared_strings = []
            if "xl/sharedStrings.xml" in z.namelist():
                ss_root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in ss_root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si"):
                    t = si.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
                    shared_strings.append(t.text if t is not None and t.text else "")

            # Sheet 1
            sheet_names = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
            if not sheet_names:
                return {}
            sheet_root = ET.fromstring(z.read(sheet_names[0]))

            rows_data = []
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            for row in sheet_root.findall(f".//{ns}row"):
                row_vals = {}
                for cell in row.findall(f"{ns}c"):
                    ref = cell.attrib.get("r", "")
                    col_letter = "".join(filter(str.isalpha, ref))
                    t_attr = cell.attrib.get("t", "")
                    v_elem = cell.find(f"{ns}v")
                    val = v_elem.text if v_elem is not None else ""
                    if t_attr == "s" and val.isdigit():
                        idx = int(val)
                        val = shared_strings[idx] if idx < len(shared_strings) else val
                    row_vals[col_letter] = val
                rows_data.append(row_vals)

            if not rows_data:
                return {}

            header_row = rows_data[0]
            col_map = {}
            for col_letter, text in header_row.items():
                low = str(text).strip().lower()
                if "provider" in low:
                    col_map["provider"] = col_letter
                elif "nickname" in low:
                    col_map["nickname"] = col_letter
                elif "api key" in low or "key" in low:
                    col_map["key"] = col_letter

            prov_c = col_map.get("provider", "A")
            nick_c = col_map.get("nickname", "B")
            key_c = col_map.get("key", "C")

            keys_data = {}
            for r in rows_data[1:]:
                nick = str(r.get(nick_c, "")).strip()
                prov = str(r.get(prov_c, "")).strip()
                key = str(r.get(key_c, "")).strip()
                if nick and key:
                    if key.lower().startswith("bearer "):
                        key = key[7:].strip()
                    keys_data[nick] = {"provider": prov, "key": key}

            return keys_data
    except Exception:
        return {}


def save_provider_to_json(provider_name: str, base_url: str, model: str) -> Tuple[bool, str]:
    """Saves or updates a provider and model in providers.json."""
    if not provider_name.strip():
        return (False, "Provider name cannot be empty.")
    try:
        data = load_providers_json()
        prov_key = provider_name.strip()
        if prov_key not in data:
            data[prov_key] = {
                "base_url": base_url.strip(),
                "models": [model.strip()] if model.strip() else []
            }
        else:
            if base_url.strip():
                data[prov_key]["base_url"] = base_url.strip()
            models_list = data[prov_key].get("models", [])
            if not isinstance(models_list, list):
                models_list = []
            if model.strip() and model.strip() not in models_list:
                models_list.append(model.strip())
            data[prov_key]["models"] = models_list

        PROVIDERS_FILE.write_text(
            json.dumps({"providers": data}, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8"
        )
        return (True, f"Successfully saved '{model}' under provider '{prov_key}' to providers.json!")
    except Exception as exc:
        return (False, f"Error saving to providers.json: {exc}")


# =============================================================================
# MODERN DARK THEME STYLESHEET (Matches Switcher_v3.py & fixes combo popups)
# =============================================================================

API_TESTER_STYLESHEET = """
QWidget {
    background-color: #12131c;
    color: #e2e8f0;
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    font-size: 13px;
}

QMainWindow, QDialog {
    background-color: #12131c;
}

/* Card Containers & Group Boxes */
QGroupBox {
    font-weight: bold;
    border: 1px solid #2d3148;
    border-radius: 9px;
    margin-top: 10px;
    padding-top: 14px;
    background-color: #1a1c2b;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    color: #94a3b8;
    font-size: 12px;
}

/* Labels */
QLabel {
    color: #cbd5e1;
    font-size: 13px;
}

/* Form Inputs & Comboboxes */
QLineEdit, QComboBox, QSpinBox {
    background-color: #222638;
    border: 1px solid #363b57;
    border-radius: 6px;
    padding: 6px 10px;
    color: #f8fafc;
    font-size: 13px;
    min-height: 22px;
}

QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
    border: 1.5px solid #6366f1;
    background-color: #272c40;
}

/* Editable Combobox Support */
QComboBox:editable {
    background-color: #222638;
    color: #f8fafc;
}

QComboBox:editable:focus {
    border: 1.5px solid #6366f1;
    background-color: #272c40;
}

QComboBox QLineEdit {
    background-color: transparent;
    color: #f8fafc;
    border: none;
    padding: 0;
    font-size: 13px;
}

/* Combobox Arrow */
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    border-left: 1px solid #363b57;
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
}

QComboBox::down-arrow {
    width: 0;
    height: 0;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 6px solid #94a3b8;
}

QComboBox::down-arrow:hover {
    border-top-color: #6366f1;
}

/* Combobox Dropdown Popups (QAbstractItemView & QListView) - Guaranteed High Visibility */
QComboBox QAbstractItemView, QListView {
    background-color: #1a1c2b;
    color: #f8fafc;
    border: 1px solid #3b4263;
    border-radius: 8px;
    padding: 4px;
    selection-background-color: #6366f1;
    selection-color: #ffffff;
    outline: none;
}

QComboBox QAbstractItemView::item, QListView::item {
    min-height: 28px;
    padding: 6px 10px;
    border-radius: 4px;
    color: #f8fafc;
}

QComboBox QAbstractItemView::item:hover, QListView::item:hover {
    background-color: #2b3047;
    color: #ffffff;
}

QComboBox QAbstractItemView::item:selected, QListView::item:selected {
    background-color: #6366f1;
    color: #ffffff;
}

/* Action Buttons */
QPushButton {
    background-color: #222638;
    border: 1px solid #363b57;
    border-radius: 6px;
    padding: 7px 14px;
    color: #cbd5e1;
    font-weight: 500;
}

QPushButton:hover {
    background-color: #2c324a;
    border-color: #4b5275;
    color: #f8fafc;
}

QPushButton:pressed {
    background-color: #1a1e2e;
}

QPushButton:disabled {
    background-color: #181a26;
    border-color: #25283b;
    color: #64748b;
}

QPushButton#testBtn {
    background-color: #6366f1;
    border: 1px solid #6366f1;
    color: #ffffff;
    font-weight: bold;
    font-size: 13px;
    padding: 8px 22px;
}

QPushButton#testBtn:hover {
    background-color: #4f46e5;
    border-color: #4f46e5;
}

QPushButton#testBtn:disabled {
    background-color: #312e81;
    border-color: #312e81;
    color: #94a3b8;
}

/* Text Editors */
QTextEdit {
    background-color: #161824;
    border: 1px solid #2d3148;
    border-radius: 6px;
    color: #f8fafc;
    padding: 10px;
    font-size: 12px;
}

/* Tabs & Tab Bars */
QTabWidget::pane {
    border: 1px solid #2d3148;
    border-radius: 6px;
    background-color: #1a1c2b;
}

QTabBar::tab {
    background-color: #161824;
    color: #94a3b8;
    border: 1px solid #2d3148;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 8px 16px;
    margin-right: 2px;
}

QTabBar::tab:selected {
    background-color: #1a1c2b;
    color: #f8fafc;
    font-weight: bold;
    border-bottom: 2px solid #6366f1;
}

QTabBar::tab:hover:!selected {
    background-color: #222638;
    color: #cbd5e1;
}

/* Tables */
QTableWidget {
    background-color: #161824;
    border: 1px solid #2d3148;
    border-radius: 6px;
    color: #f8fafc;
    gridline-color: #2d3148;
}

QTableWidget QHeaderView::section {
    background-color: #1a1c2b;
    color: #94a3b8;
    padding: 6px;
    border: 1px solid #2d3148;
    font-weight: bold;
}

QTableWidget::item:selected {
    background-color: #6366f1;
    color: #ffffff;
}

/* Scrollbars */
QScrollBar:vertical {
    background-color: #161824;
    width: 8px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #363b57;
    min-height: 20px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background-color: #6366f1;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


# =============================================================================
# PYSIDE6 IMPLEMENTATION (Primary GUI)
# =============================================================================

def run_pyside6_gui():
    from PySide6.QtCore import Qt, QThread, Signal
    from PySide6.QtWidgets import (
        QApplication,
        QMainWindow,
        QWidget,
        QVBoxLayout,
        QHBoxLayout,
        QFormLayout,
        QGridLayout,
        QLabel,
        QLineEdit,
        QComboBox,
        QPushButton,
        QTextEdit,
        QTabWidget,
        QProgressBar,
        QGroupBox,
        QCheckBox,
        QSpinBox,
        QTableWidget,
        QTableWidgetItem,
        QHeaderView,
        QMessageBox,
        QInputDialog,
        QListView,
    )
    from PySide6.QtGui import QFont, QColor

    class ApiWorkerThread(QThread):
        """Worker thread for non-blocking HTTP verification."""
        test_finished = Signal(object)

        def __init__(self, kwargs):
            super().__init__()
            self.kwargs = kwargs

        def run(self):
            result = run_api_test(**self.kwargs)
            self.test_finished.emit(result)

    class ApiTesterWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.providers = load_providers_json()
            self.saved_keys = load_keys_from_excel()
            self.worker_thread: Optional[ApiWorkerThread] = None
            self.last_result: Optional[ApiTestResult] = None
            self.history: List[Tuple[str, str, str, str, str, str]] = []

            self.setWindowTitle("Key Switcher - API Connection Tester")
            self.resize(780, 660)
            self.setMinimumSize(640, 480)

            self.setup_ui()
            self.apply_theme()
            self.populate_presets()

        def setup_ui(self):
            central = QWidget()
            self.setCentralWidget(central)
            main_layout = QVBoxLayout(central)
            main_layout.setContentsMargins(18, 16, 18, 16)
            main_layout.setSpacing(12)

            # -----------------------------------------------------------------
            # 1. Header
            # -----------------------------------------------------------------
            header_layout = QHBoxLayout()
            title_vbox = QVBoxLayout()
            title_lbl = QLabel("API Connection Tester")
            title_lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #818cf8;")
            subtitle_lbl = QLabel("Test Base URL, API Key, and Model connectivity in real time")
            subtitle_lbl.setStyleSheet("font-size: 12px; color: #94a3b8;")
            title_vbox.addWidget(title_lbl)
            title_vbox.addWidget(subtitle_lbl)
            header_layout.addLayout(title_vbox)
            header_layout.addStretch()

            self.badge_lbl = QLabel("READY")
            self.badge_lbl.setStyleSheet(
                "background-color: #272c40; color: #94a3b8; font-weight: bold; "
                "padding: 4px 14px; border-radius: 12px; font-size: 11px;"
            )
            header_layout.addWidget(self.badge_lbl)
            main_layout.addLayout(header_layout)

            # -----------------------------------------------------------------
            # 2. Presets Group (Quick Load)
            # -----------------------------------------------------------------
            preset_group = QGroupBox("Presets & Auto-Fill")
            preset_layout = QGridLayout(preset_group)
            preset_layout.setContentsMargins(14, 12, 14, 12)
            preset_layout.setSpacing(10)

            preset_layout.addWidget(QLabel("Provider Preset:"), 0, 0)
            self.provider_combo = QComboBox()
            self.provider_combo.setView(QListView())
            self.provider_combo.setMaxVisibleItems(12)
            self.provider_combo.currentIndexChanged.connect(self.on_provider_selected)
            preset_layout.addWidget(self.provider_combo, 0, 1)

            preset_layout.addWidget(QLabel("Saved Key (keys.xlsx):"), 0, 2)
            self.saved_key_combo = QComboBox()
            self.saved_key_combo.setView(QListView())
            self.saved_key_combo.setMaxVisibleItems(12)
            self.saved_key_combo.currentIndexChanged.connect(self.on_saved_key_selected)
            preset_layout.addWidget(self.saved_key_combo, 0, 3)

            main_layout.addWidget(preset_group)

            # -----------------------------------------------------------------
            # 3. Connection Configuration Form
            # -----------------------------------------------------------------
            form_group = QGroupBox("API Configuration")
            form_layout = QFormLayout(form_group)
            form_layout.setContentsMargins(14, 12, 14, 12)
            form_layout.setSpacing(10)

            # Base URL
            self.url_edit = QLineEdit()
            self.url_edit.setPlaceholderText("e.g. https://openrouter.ai/api/v1 or https://api.groq.com/openai/v1")
            form_layout.addRow("Base URL *:", self.url_edit)

            # API Key
            key_layout = QHBoxLayout()
            self.key_edit = QLineEdit()
            self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
            self.key_edit.setPlaceholderText("Enter API Key (optional for local Ollama)")
            self.toggle_key_btn = QPushButton("👁 Show")
            self.toggle_key_btn.setFixedWidth(74)
            self.toggle_key_btn.clicked.connect(self.toggle_key_visibility)
            key_layout.addWidget(self.key_edit)
            key_layout.addWidget(self.toggle_key_btn)
            form_layout.addRow("API Key:", key_layout)

            # Model (Editable combo box)
            self.model_combo = QComboBox()
            self.model_combo.setEditable(True)
            self.model_combo.setView(QListView())
            self.model_combo.setMaxVisibleItems(12)
            self.model_combo.setPlaceholderText("e.g. openai/gpt-oss-120b or llama-3.3-70b-versatile")
            form_layout.addRow("Model *:", self.model_combo)

            # Protocol
            self.protocol_combo = QComboBox()
            self.protocol_combo.setView(QListView())
            self.protocol_combo.setMaxVisibleItems(12)
            self.protocol_combo.addItem("Auto-Detect (OpenAI & Anthropic)", "auto")
            self.protocol_combo.addItem("OpenAI Compatible (/chat/completions)", "openai")
            self.protocol_combo.addItem("Anthropic Compatible (/messages)", "anthropic")
            self.protocol_combo.addItem("Direct Endpoint URL (Use exact Base URL)", "direct")
            form_layout.addRow("Protocol:", self.protocol_combo)

            main_layout.addWidget(form_group)

            # -----------------------------------------------------------------
            # 4. Action Buttons & Progress
            # -----------------------------------------------------------------
            btn_layout = QHBoxLayout()
            self.test_btn = QPushButton("🚀 TEST API")
            self.test_btn.setObjectName("testBtn")
            self.test_btn.setMinimumHeight(40)
            self.test_btn.clicked.connect(self.start_test)
            btn_layout.addWidget(self.test_btn)

            self.cancel_btn = QPushButton("Cancel")
            self.cancel_btn.setMinimumHeight(40)
            self.cancel_btn.setEnabled(False)
            self.cancel_btn.clicked.connect(self.cancel_test)
            btn_layout.addWidget(self.cancel_btn)

            self.copy_btn = QPushButton("📋 Copy Report")
            self.copy_btn.setMinimumHeight(40)
            self.copy_btn.setEnabled(False)
            self.copy_btn.clicked.connect(self.copy_report)
            btn_layout.addWidget(self.copy_btn)

            self.save_btn = QPushButton("💾 Save Preset")
            self.save_btn.setMinimumHeight(40)
            self.save_btn.setToolTip("Save this Base URL and Model to providers.json")
            self.save_btn.clicked.connect(self.save_to_providers)
            btn_layout.addWidget(self.save_btn)

            self.clear_btn = QPushButton("Clear")
            self.clear_btn.setMinimumHeight(40)
            self.clear_btn.clicked.connect(self.clear_results)
            btn_layout.addWidget(self.clear_btn)

            main_layout.addLayout(btn_layout)

            self.progress_bar = QProgressBar()
            self.progress_bar.setRange(0, 0)
            self.progress_bar.setFixedHeight(6)
            self.progress_bar.hide()
            main_layout.addWidget(self.progress_bar)

            # -----------------------------------------------------------------
            # 5. Status Banner (Prominent Working/Failed Card)
            # -----------------------------------------------------------------
            self.status_banner = QLabel("Ready to test. Enter connection details and click 'TEST API'.")
            self.status_banner.setWordWrap(True)
            self.status_banner.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.set_banner_idle()
            main_layout.addWidget(self.status_banner)

            # -----------------------------------------------------------------
            # 6. Results Tabs
            # -----------------------------------------------------------------
            self.tabs = QTabWidget()

            # Tab 1: Overview
            tab_overview = QWidget()
            tab_over_layout = QVBoxLayout(tab_overview)
            tab_over_layout.setContentsMargins(10, 10, 10, 10)

            meta_grid = QGridLayout()
            meta_grid.addWidget(QLabel("<b>Status:</b>"), 0, 0)
            self.meta_status = QLabel("-")
            meta_grid.addWidget(self.meta_status, 0, 1)

            meta_grid.addWidget(QLabel("<b>Latency:</b>"), 0, 2)
            self.meta_latency = QLabel("-")
            meta_grid.addWidget(self.meta_latency, 0, 3)

            meta_grid.addWidget(QLabel("<b>Endpoint:</b>"), 1, 0)
            self.meta_endpoint = QLabel("-")
            self.meta_endpoint.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            meta_grid.addWidget(self.meta_endpoint, 1, 1)

            meta_grid.addWidget(QLabel("<b>Tokens:</b>"), 1, 2)
            self.meta_tokens = QLabel("-")
            meta_grid.addWidget(self.meta_tokens, 1, 3)

            tab_over_layout.addLayout(meta_grid)

            tab_over_layout.addWidget(QLabel("<b>Model Response / Diagnosis:</b>"))
            self.response_text = QTextEdit()
            self.response_text.setReadOnly(True)
            self.response_text.setPlaceholderText("Model output or error diagnosis will appear here...")
            tab_over_layout.addWidget(self.response_text)

            self.tabs.addTab(tab_overview, "Overview")

            # Tab 2: Raw Response & Headers
            tab_raw = QWidget()
            tab_raw_layout = QVBoxLayout(tab_raw)
            tab_raw_layout.setContentsMargins(10, 10, 10, 10)
            self.raw_text = QTextEdit()
            self.raw_text.setReadOnly(True)
            self.raw_text.setFont(QFont("Consolas", 10))
            tab_raw_layout.addWidget(self.raw_text)
            self.tabs.addTab(tab_raw, "Raw Response")

            # Tab 3: History
            tab_hist = QWidget()
            tab_hist_layout = QVBoxLayout(tab_hist)
            tab_hist_layout.setContentsMargins(10, 10, 10, 10)
            self.history_table = QTableWidget(0, 5)
            self.history_table.setHorizontalHeaderLabels(["Time", "Model", "Result", "HTTP Code", "Latency"])
            self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            tab_hist_layout.addWidget(self.history_table)
            self.tabs.addTab(tab_hist, "History")

            main_layout.addWidget(self.tabs)

        def apply_theme(self):
            app_inst = QApplication.instance()
            if app_inst:
                app_inst.setStyleSheet(API_TESTER_STYLESHEET)
            else:
                self.setStyleSheet(API_TESTER_STYLESHEET)

        def set_banner_idle(self):
            self.status_banner.setText("Ready to test. Enter connection details and click 'TEST API'.")
            self.status_banner.setStyleSheet(
                "background-color: #1e2235; color: #94a3b8; border: 1px solid #2d3148; "
                "border-radius: 8px; padding: 12px; font-size: 13px; font-weight: bold;"
            )
            self.badge_lbl.setText("READY")
            self.badge_lbl.setStyleSheet(
                "background-color: #272c40; color: #94a3b8; font-weight: bold; "
                "padding: 4px 14px; border-radius: 12px; font-size: 11px;"
            )

        def set_banner_testing(self, target: str):
            self.status_banner.setText(f"⏳ Testing connection to: {target} ...")
            self.status_banner.setStyleSheet(
                "background-color: #1e293b; color: #60a5fa; border: 1px solid #3b82f6; "
                "border-radius: 8px; padding: 12px; font-size: 13px; font-weight: bold;"
            )
            self.badge_lbl.setText("TESTING...")
            self.badge_lbl.setStyleSheet(
                "background-color: #1d4ed8; color: #ffffff; font-weight: bold; "
                "padding: 4px 14px; border-radius: 12px; font-size: 11px;"
            )

        def set_banner_success(self, res: ApiTestResult):
            self.status_banner.setText(
                f"✔ WORKING! (HTTP {res.status_code}) — Response time: {res.latency_ms:.1f} ms"
            )
            self.status_banner.setStyleSheet(
                "background-color: #14532d; color: #bbf7d0; border: 1px solid #22c55e; "
                "border-radius: 8px; padding: 12px; font-size: 13px; font-weight: bold;"
            )
            self.badge_lbl.setText("WORKING")
            self.badge_lbl.setStyleSheet(
                "background-color: #16a34a; color: #ffffff; font-weight: bold; "
                "padding: 4px 14px; border-radius: 12px; font-size: 11px;"
            )

        def set_banner_failure(self, res: ApiTestResult):
            status_desc = f"HTTP {res.status_code}" if res.status_code else "Connection Error"
            self.status_banner.setText(
                f"✖ NOT WORKING — {status_desc}: {res.status_message}"
            )
            self.status_banner.setStyleSheet(
                "background-color: #450a0a; color: #fecaca; border: 1px solid #ef4444; "
                "border-radius: 8px; padding: 12px; font-size: 13px; font-weight: bold;"
            )
            self.badge_lbl.setText("FAILED")
            self.badge_lbl.setStyleSheet(
                "background-color: #dc2626; color: #ffffff; font-weight: bold; "
                "padding: 4px 14px; border-radius: 12px; font-size: 11px;"
            )

        def populate_presets(self):
            # 1. Populate Provider Presets
            self.provider_combo.blockSignals(True)
            self.provider_combo.addItem("Select a Preset / Provider...", "")
            for prov_name in self.providers.keys():
                self.provider_combo.addItem(prov_name, prov_name)
            self.provider_combo.addItem("Custom", "Custom")
            self.provider_combo.blockSignals(False)

            # 2. Populate Saved Keys from keys.xlsx
            self.saved_key_combo.blockSignals(True)
            self.saved_key_combo.addItem("Select from keys.xlsx...", "")
            if self.saved_keys:
                for nick, data in self.saved_keys.items():
                    prov = data.get("provider", "")
                    display = f"{nick} ({prov})" if prov else nick
                    self.saved_key_combo.addItem(display, nick)
            else:
                self.saved_key_combo.addItem("(No saved keys found)", "")
            self.saved_key_combo.blockSignals(False)

        def on_provider_selected(self, index: int):
            prov_name = self.provider_combo.currentData()
            if not prov_name or prov_name == "Custom":
                return
            prov_info = self.providers.get(prov_name, {})
            base_url = prov_info.get("base_url", "")
            models = prov_info.get("models", [])

            self.url_edit.setText(base_url)
            self.model_combo.clear()
            for m in models:
                self.model_combo.addItem(m)
            if models:
                self.model_combo.setCurrentIndex(0)

            # Auto-match saved key if one exists for this provider
            for nick, data in self.saved_keys.items():
                if data.get("provider", "").strip().lower() == prov_name.strip().lower():
                    for i in range(self.saved_key_combo.count()):
                        if self.saved_key_combo.itemData(i) == nick:
                            self.saved_key_combo.blockSignals(True)
                            self.saved_key_combo.setCurrentIndex(i)
                            self.saved_key_combo.blockSignals(False)
                            k = data.get("key", "").strip()
                            if k.lower().startswith("bearer "):
                                k = k[7:].strip()
                            self.key_edit.setText(k)
                            break
                    break

        def on_saved_key_selected(self, index: int):
            nick = self.saved_key_combo.currentData()
            if not nick or nick not in self.saved_keys:
                return
            key_info = self.saved_keys[nick]
            api_key = key_info.get("key", "").strip()
            if api_key.lower().startswith("bearer "):
                api_key = api_key[7:].strip()
            prov_name = key_info.get("provider", "").strip()

            self.key_edit.setText(api_key)

            # Match provider case-insensitively
            matched_prov = None
            for p_name in self.providers.keys():
                if p_name.strip().lower() == prov_name.lower():
                    matched_prov = p_name
                    break

            if matched_prov:
                for i in range(self.provider_combo.count()):
                    if str(self.provider_combo.itemData(i)).lower() == matched_prov.lower():
                        self.provider_combo.blockSignals(True)
                        self.provider_combo.setCurrentIndex(i)
                        self.provider_combo.blockSignals(False)
                        break

                prov_info = self.providers.get(matched_prov, {})
                base_url = prov_info.get("base_url", "")
                models = prov_info.get("models", [])

                if base_url:
                    self.url_edit.setText(base_url)

                self.model_combo.clear()
                for m in models:
                    self.model_combo.addItem(m)
                if models:
                    self.model_combo.setCurrentIndex(0)

        def toggle_key_visibility(self):
            if self.key_edit.echoMode() == QLineEdit.EchoMode.Password:
                self.key_edit.setEchoMode(QLineEdit.EchoMode.Normal)
                self.toggle_key_btn.setText("🔒 Hide")
            else:
                self.key_edit.setEchoMode(QLineEdit.EchoMode.Password)
                self.toggle_key_btn.setText("👁 Show")

        def start_test(self):
            base_url = self.url_edit.text().strip()
            model = self.model_combo.currentText().strip()
            api_key = self.key_edit.text().strip()
            protocol = self.protocol_combo.currentData() or "auto"
            prompt = DEFAULT_PROMPT
            max_tokens = DEFAULT_MAX_TOKENS
            timeout = DEFAULT_TIMEOUT
            verify_ssl = True

            if not base_url:
                QMessageBox.warning(self, "Missing URL", "Please enter a Base URL.")
                self.url_edit.setFocus()
                return
            if not model:
                QMessageBox.warning(self, "Missing Model", "Please enter or select a Model name.")
                self.model_combo.setFocus()
                return

            self.test_btn.setEnabled(False)
            self.cancel_btn.setEnabled(True)
            self.copy_btn.setEnabled(False)
            self.progress_bar.show()

            self.set_banner_testing(model or base_url)

            kwargs = {
                "base_url": base_url,
                "api_key": api_key,
                "model": model,
                "protocol": protocol,
                "prompt": prompt,
                "max_tokens": max_tokens,
                "timeout": timeout,
                "verify_ssl": verify_ssl,
            }

            self.worker_thread = ApiWorkerThread(kwargs)
            self.worker_thread.test_finished.connect(self.on_test_finished)
            self.worker_thread.start()

        def cancel_test(self):
            if self.worker_thread and self.worker_thread.isRunning():
                self.worker_thread.terminate()
                self.worker_thread.wait()
                self.set_banner_idle()
                self.status_banner.setText("Test cancelled by user.")
                self.test_btn.setEnabled(True)
                self.cancel_btn.setEnabled(False)
                self.progress_bar.hide()

        def on_test_finished(self, result: ApiTestResult):
            self.last_result = result
            self.test_btn.setEnabled(True)
            self.cancel_btn.setEnabled(False)
            self.copy_btn.setEnabled(True)
            self.progress_bar.hide()

            # 1. Update Banner
            if result.is_working:
                self.set_banner_success(result)
            else:
                self.set_banner_failure(result)

            # 2. Update Overview Meta
            status_text = f"{result.status_code} ({result.status_message})" if result.status_code else result.status_message
            self.meta_status.setText(status_text)
            self.meta_latency.setText(f"{result.latency_ms:.1f} ms")
            self.meta_endpoint.setText(result.endpoint_used or "-")
            self.meta_tokens.setText(result.token_usage or "-")

            # 3. Response / Diagnosis Text Box
            content_parts = []
            if result.is_working:
                content_parts.append(f"=== MODEL RESPONSE ===\n{result.model_reply}\n")
                if result.troubleshooting:
                    content_parts.append(f"=== NOTE ===\n{result.troubleshooting}")
            else:
                content_parts.append(f"=== ERROR DIAGNOSIS ===\n{result.error_details}\n")
                if result.troubleshooting:
                    content_parts.append(f"=== HOW TO FIX ===\n{result.troubleshooting}")

            self.response_text.setPlainText("\n".join(content_parts))

            # 4. Raw Response Tab
            raw_view = [f"HTTP Status: {result.status_code} {result.status_message}"]
            raw_view.append(f"Endpoint: {result.endpoint_used}")
            raw_view.append(f"Protocol Style: {result.protocol}")
            raw_view.append("\n--- Response Headers ---")
            for hk, hv in result.response_headers.items():
                raw_view.append(f"{hk}: {hv}")
            raw_view.append("\n--- Response Body ---")
            raw_view.append(result.raw_response or "(No response body)")
            self.raw_text.setPlainText("\n".join(raw_view))

            # 5. History Table
            row_pos = self.history_table.rowCount()
            self.history_table.insertRow(row_pos)
            res_str = "✔ Working" if result.is_working else "✖ Failed"
            self.history_table.setItem(row_pos, 0, QTableWidgetItem(result.timestamp.split(" ")[1]))
            self.history_table.setItem(row_pos, 1, QTableWidgetItem(self.model_combo.currentText()))
            self.history_table.setItem(row_pos, 2, QTableWidgetItem(res_str))
            self.history_table.setItem(row_pos, 3, QTableWidgetItem(str(result.status_code)))
            self.history_table.setItem(row_pos, 4, QTableWidgetItem(f"{result.latency_ms:.0f} ms"))

        def copy_report(self):
            if not self.last_result:
                return
            report = self.last_result.to_markdown_report(
                base_url=self.url_edit.text(),
                model=self.model_combo.currentText(),
            )
            clipboard = QApplication.clipboard()
            clipboard.setText(report)
            QMessageBox.information(self, "Copied", "Diagnostic report copied to clipboard!")

        def save_to_providers(self):
            base_url = self.url_edit.text().strip()
            model = self.model_combo.currentText().strip()
            if not base_url:
                QMessageBox.warning(self, "Incomplete", "Please specify a Base URL.")
                return
            if not model:
                QMessageBox.warning(self, "Incomplete", "Please specify a Model name.")
                return

            current_prov = self.provider_combo.currentData() or ""
            if not current_prov or current_prov == "Custom":
                prov_name, ok = QInputDialog.getText(
                    self, "Provider Name", "Enter provider name for providers.json:"
                )
                if not ok or not prov_name.strip():
                    return
                current_prov = prov_name.strip()

            success, msg = save_provider_to_json(current_prov, base_url, model)
            if success:
                QMessageBox.information(self, "Saved", msg)
                self.providers = load_providers_json()
                self.populate_presets()
                # Select the saved provider
                for i in range(self.provider_combo.count()):
                    if self.provider_combo.itemData(i) == current_prov:
                        self.provider_combo.setCurrentIndex(i)
                        break
            else:
                QMessageBox.critical(self, "Error", msg)

        def clear_results(self):
            self.set_banner_idle()
            self.meta_status.setText("-")
            self.meta_latency.setText("-")
            self.meta_endpoint.setText("-")
            self.meta_tokens.setText("-")
            self.response_text.clear()
            self.raw_text.clear()
            self.last_result = None
            self.copy_btn.setEnabled(False)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(API_TESTER_STYLESHEET)
    window = ApiTesterWindow()
    window.show()
    sys.exit(app.exec())


# =============================================================================
# TKINTER IMPLEMENTATION (Fallback GUI - Zero External Dependencies)
# =============================================================================

def run_tkinter_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox
    import threading

    class TkinterApiTester(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Key Switcher - API Connection Tester (Standard GUI)")
            self.geometry("740, 640")
            self.minsize(600, 480)

            self.providers = load_providers_json()
            self.saved_keys = load_keys_from_excel()
            self.last_result: Optional[ApiTestResult] = None
            self.is_running = False

            self.setup_ui()
            self.populate_presets()

        def setup_ui(self):
            main_frame = ttk.Frame(self, padding=12)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Header
            header_frame = ttk.Frame(main_frame)
            header_frame.pack(fill=tk.X, pady=(0, 10))
            ttk.Label(header_frame, text="API Connection Tester", font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT)
            self.badge_var = tk.StringVar(value="READY")
            self.badge_lbl = tk.Label(header_frame, textvariable=self.badge_var, bg="#e2e8f0", fg="#475569", font=("Segoe UI", 9, "bold"), padx=8, pady=2)
            self.badge_lbl.pack(side=tk.RIGHT)

            # Presets Frame
            preset_lf = ttk.LabelFrame(main_frame, text="Presets & Auto-Fill", padding=8)
            preset_lf.pack(fill=tk.X, pady=(0, 8))

            ttk.Label(preset_lf, text="Provider:").grid(row=0, column=0, sticky="w", padx=4, pady=2)
            self.provider_var = tk.StringVar()
            self.provider_combo = ttk.Combobox(preset_lf, textvariable=self.provider_var, state="readonly", width=22)
            self.provider_combo.grid(row=0, column=1, sticky="w", padx=4, pady=2)
            self.provider_combo.bind("<<ComboboxSelected>>", self.on_provider_selected)

            ttk.Label(preset_lf, text="Saved Key:").grid(row=0, column=2, sticky="w", padx=(16, 4), pady=2)
            self.key_select_var = tk.StringVar()
            self.key_combo = ttk.Combobox(preset_lf, textvariable=self.key_select_var, state="readonly", width=26)
            self.key_combo.grid(row=0, column=3, sticky="w", padx=4, pady=2)
            self.key_combo.bind("<<ComboboxSelected>>", self.on_saved_key_selected)

            # Config Frame
            cfg_lf = ttk.LabelFrame(main_frame, text="API Configuration", padding=8)
            cfg_lf.pack(fill=tk.X, pady=(0, 8))

            ttk.Label(cfg_lf, text="Base URL:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
            self.url_var = tk.StringVar()
            self.url_entry = ttk.Entry(cfg_lf, textvariable=self.url_var, width=54)
            self.url_entry.grid(row=0, column=1, columnspan=2, sticky="we", padx=4, pady=4)

            ttk.Label(cfg_lf, text="API Key:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
            self.key_var = tk.StringVar()
            self.key_entry = ttk.Entry(cfg_lf, textvariable=self.key_var, show="•", width=42)
            self.key_entry.grid(row=1, column=1, sticky="we", padx=4, pady=4)
            self.toggle_btn = ttk.Button(cfg_lf, text="Show", width=8, command=self.toggle_key)
            self.toggle_btn.grid(row=1, column=2, sticky="w", padx=4, pady=4)

            ttk.Label(cfg_lf, text="Model:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
            self.model_var = tk.StringVar()
            self.model_combo = ttk.Combobox(cfg_lf, textvariable=self.model_var, width=42)
            self.model_combo.grid(row=2, column=1, sticky="we", padx=4, pady=4)

            ttk.Label(cfg_lf, text="Protocol:").grid(row=3, column=0, sticky="w", padx=4, pady=4)
            self.protocol_var = tk.StringVar(value="auto")
            proto_combo = ttk.Combobox(cfg_lf, textvariable=self.protocol_var, values=["auto", "openai", "anthropic", "direct"], state="readonly", width=22)
            proto_combo.grid(row=3, column=1, sticky="w", padx=4, pady=4)

            cfg_lf.columnconfigure(1, weight=1)

            # Action Buttons
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, pady=(0, 8))

            self.test_btn = ttk.Button(btn_frame, text="🚀 TEST API", command=self.start_test)
            self.test_btn.pack(side=tk.LEFT, padx=(0, 6), ipadx=10, ipady=4)

            self.copy_btn = ttk.Button(btn_frame, text="📋 Copy Report", command=self.copy_report, state="disabled")
            self.copy_btn.pack(side=tk.LEFT, padx=6)

            self.save_btn = ttk.Button(btn_frame, text="💾 Save Preset", command=self.save_to_providers)
            self.save_btn.pack(side=tk.LEFT, padx=6)

            self.clear_btn = ttk.Button(btn_frame, text="Clear", command=self.clear_results)
            self.clear_btn.pack(side=tk.LEFT, padx=6)

            # Status Banner
            self.banner_var = tk.StringVar(value="Ready to test. Enter connection details and click 'TEST API'.")
            self.banner_lbl = tk.Label(main_frame, textvariable=self.banner_var, bg="#f1f5f9", fg="#475569", font=("Segoe UI", 10, "bold"), relief="solid", bd=1, padx=8, pady=8)
            self.banner_lbl.pack(fill=tk.X, pady=(0, 8))

            # Notebook Tabs
            notebook = ttk.Notebook(main_frame)
            notebook.pack(fill=tk.BOTH, expand=True)

            # Tab 1: Output
            tab_out = ttk.Frame(notebook, padding=6)
            notebook.add(tab_out, text="Overview & Response")

            self.out_text = tk.Text(tab_out, wrap="word", height=12, font=("Segoe UI", 10))
            self.out_text.pack(fill=tk.BOTH, expand=True)

            # Tab 2: Raw
            tab_raw = ttk.Frame(notebook, padding=6)
            notebook.add(tab_raw, text="Raw Response")
            self.raw_text = tk.Text(tab_raw, wrap="none", height=12, font=("Consolas", 9))
            self.raw_text.pack(fill=tk.BOTH, expand=True)

        def populate_presets(self):
            prov_names = ["(Select Preset)"] + list(self.providers.keys()) + ["Custom"]
            self.provider_combo["values"] = prov_names
            self.provider_combo.current(0)

            key_names = ["(Select Saved Key)"] + [f"{k} ({v.get('provider','')})" for k, v in self.saved_keys.items()]
            self.key_combo["values"] = key_names
            self.key_combo.current(0)

        def on_provider_selected(self, event=None):
            sel = self.provider_var.get()
            if not sel or sel in ("(Select Preset)", "Custom"):
                return
            if sel in self.providers:
                info = self.providers[sel]
                self.url_var.set(info.get("base_url", ""))
                models = info.get("models", [])
                self.model_combo["values"] = models
                if models:
                    self.model_var.set(models[0])

            # Auto-match saved key if one exists for this provider
            for nick, data in self.saved_keys.items():
                if data.get("provider", "").strip().lower() == sel.strip().lower():
                    k = data.get("key", "").strip()
                    if k.lower().startswith("bearer "):
                        k = k[7:].strip()
                    self.key_var.set(k)
                    for item in self.key_combo["values"]:
                        if item.startswith(nick):
                            self.key_select_var.set(item)
                            break
                    break

        def on_saved_key_selected(self, event=None):
            sel = self.key_select_var.get()
            nick = sel.split(" (")[0]
            if nick in self.saved_keys:
                data = self.saved_keys[nick]
                k = data.get("key", "").strip()
                if k.lower().startswith("bearer "):
                    k = k[7:].strip()
                self.key_var.set(k)
                prov = data.get("provider", "").strip()

                # Case-insensitive provider match
                matched = None
                for p in self.providers.keys():
                    if p.strip().lower() == prov.lower():
                        matched = p
                        break
                if matched:
                    self.provider_var.set(matched)
                    info = self.providers[matched]
                    self.url_var.set(info.get("base_url", ""))
                    models = info.get("models", [])
                    self.model_combo["values"] = models
                    if models:
                        self.model_var.set(models[0])

        def toggle_key(self):
            if self.key_entry.cget("show") == "":
                self.key_entry.config(show="•")
                self.toggle_btn.config(text="Show")
            else:
                self.key_entry.config(show="")
                self.toggle_btn.config(text="Hide")

        def start_test(self):
            if self.is_running:
                return
            base_url = self.url_var.get().strip()
            model = self.model_var.get().strip()
            if not base_url:
                messagebox.showwarning("Missing URL", "Please enter a Base URL.")
                return
            if not model:
                messagebox.showwarning("Missing Model", "Please enter a Model name.")
                return

            self.is_running = True
            self.test_btn.config(state="disabled")
            self.banner_var.set(f"⏳ Testing connection to: {model or base_url} ...")
            self.banner_lbl.config(bg="#eff6ff", fg="#1e40af")
            self.badge_var.set("TESTING...")
            self.badge_lbl.config(bg="#bfdbfe", fg="#1e3a8a")

            args = {
                "base_url": base_url,
                "api_key": self.key_var.get().strip(),
                "model": model,
                "protocol": self.protocol_var.get(),
            }

            def worker():
                res = run_api_test(**args)
                self.after(0, self.on_test_finished, res)

            threading.Thread(target=worker, daemon=True).start()

        def on_test_finished(self, res: ApiTestResult):
            self.is_running = False
            self.test_btn.config(state="normal")
            self.copy_btn.config(state="normal")
            self.last_result = res

            if res.is_working:
                self.banner_var.set(f"✔ WORKING! (HTTP {res.status_code}) — {res.latency_ms:.1f} ms")
                self.banner_lbl.config(bg="#dcfce7", fg="#14532d")
                self.badge_var.set("WORKING")
                self.badge_lbl.config(bg="#22c55e", fg="white")
            else:
                self.banner_var.set(f"✖ NOT WORKING — HTTP {res.status_code}: {res.status_message}")
                self.banner_lbl.config(bg="#fee2e2", fg="#7f1d1d")
                self.badge_var.set("FAILED")
                self.badge_lbl.config(bg="#ef4444", fg="white")

            out = []
            if res.is_working:
                out.append(f"STATUS: WORKING (HTTP {res.status_code})")
                out.append(f"LATENCY: {res.latency_ms:.1f} ms")
                out.append(f"ENDPOINT: {res.endpoint_used}")
                if res.token_usage:
                    out.append(f"TOKENS: {res.token_usage}")
                out.append(f"\n--- MODEL RESPONSE ---\n{res.model_reply}")
            else:
                out.append(f"STATUS: FAILED (HTTP {res.status_code} {res.status_message})")
                out.append(f"ENDPOINT: {res.endpoint_used}")
                out.append(f"ERROR: {res.error_details}")
                if res.troubleshooting:
                    out.append(f"\n--- HOW TO FIX ---\n{res.troubleshooting}")

            self.out_text.delete("1.0", tk.END)
            self.out_text.insert(tk.END, "\n".join(out))

            self.raw_text.delete("1.0", tk.END)
            self.raw_text.insert(tk.END, res.raw_response or "(No response)")

        def copy_report(self):
            if not self.last_result:
                return
            report = self.last_result.to_markdown_report(
                base_url=self.url_var.get(),
                model=self.model_var.get(),
            )
            self.clipboard_clear()
            self.clipboard_append(report)
            messagebox.showinfo("Copied", "Diagnostic report copied to clipboard!")

        def save_to_providers(self):
            from tkinter import simpledialog
            base_url = self.url_var.get().strip()
            model = self.model_var.get().strip()
            if not base_url:
                messagebox.showwarning("Incomplete", "Please enter a Base URL.")
                return
            if not model:
                messagebox.showwarning("Incomplete", "Please enter a Model name.")
                return

            current_prov = self.provider_var.get()
            if not current_prov or current_prov in ("(Select Preset)", "Custom"):
                prov_name = simpledialog.askstring("Provider Name", "Enter provider name for providers.json:")
                if not prov_name or not prov_name.strip():
                    return
                current_prov = prov_name.strip()

            success, msg = save_provider_to_json(current_prov, base_url, model)
            if success:
                messagebox.showinfo("Saved", msg)
                self.providers = load_providers_json()
                self.populate_presets()
            else:
                messagebox.showerror("Error", msg)

        def clear_results(self):
            self.banner_var.set("Ready to test. Enter connection details and click 'TEST API'.")
            self.banner_lbl.config(bg="#f1f5f9", fg="#475569")
            self.badge_var.set("READY")
            self.badge_lbl.config(bg="#e2e8f0", fg="#475569")
            self.out_text.delete("1.0", tk.END)
            self.raw_text.delete("1.0", tk.END)
            self.copy_btn.config(state="disabled")

    app = TkinterApiTester()
    app.mainloop()


# =============================================================================
# CLI / ENTRY POINT
# =============================================================================

def main():
    """Tries PySide6 first; if unavailable, seamlessly launches Tkinter GUI."""
    try:
        import PySide6
        run_pyside6_gui()
    except ImportError:
        print("[Notice] PySide6 not found in this Python environment. Launching standard Tkinter GUI...")
        run_tkinter_gui()


if __name__ == "__main__":
    main()

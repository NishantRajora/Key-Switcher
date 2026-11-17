"""
test_ccgate.py — Comprehensive Unit & Integration Tests for ccgate.py
"""

import io
import json
import unittest
from ccgate import (
    build_chat_url,
    translate_messages_to_openai,
    translate_tools_to_openai,
    translate_tool_choice_to_openai,
    build_openai_request,
    translate_openai_response,
    StreamTranslator,
    build_error_response,
    GatewayHandler,
    load_providers_json,
    load_keys_from_excel,
    find_key_for_provider,
)


class TestCcgate(unittest.TestCase):

    def test_build_chat_url(self):
        self.assertEqual(
            build_chat_url("https://api.groq.com/openai/v1"),
            "https://api.groq.com/openai/v1/chat/completions"
        )
        self.assertEqual(
            build_chat_url("https://openrouter.ai/api"),
            "https://openrouter.ai/api/v1/chat/completions"
        )
        self.assertEqual(
            build_chat_url("https://integrate.api.nvidia.com/v1/chat/completions"),
            "https://integrate.api.nvidia.com/v1/chat/completions"
        )

    def test_translate_text_messages(self):
        msgs = [
            {"role": "user", "content": "Hello"},
            {"role": "assistant", "content": "Hi there!"},
        ]
        result = translate_messages_to_openai(msgs)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0], {"role": "user", "content": "Hello"})
        self.assertEqual(result[1], {"role": "assistant", "content": "Hi there!"})

    def test_translate_thinking_blocks(self):
        # When assistant message has only thinking block
        msgs = [
            {"role": "assistant", "content": [
                {"type": "thinking", "thinking": "Let me think...", "signature": "xyz"}
            ]}
        ]
        result = translate_messages_to_openai(msgs)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["role"], "assistant")
        self.assertEqual(result[0]["content"], "")  # Must not be None!

    def test_translate_tool_use_and_result(self):
        # Assistant invokes tool
        msgs = [
            {"role": "assistant", "content": [
                {"type": "text", "text": "Checking files..."},
                {"type": "tool_use", "id": "call_123", "name": "ls", "input": {"path": "."}}
            ]},
            {"role": "user", "content": [
                {"type": "tool_result", "tool_use_id": "call_123", "content": "file1.txt\nfile2.txt"},
                {"type": "text", "text": "Now delete file2.txt"}
            ]}
        ]
        result = translate_messages_to_openai(msgs)
        self.assertEqual(len(result), 3)

        # 1: Assistant with tool_calls
        self.assertEqual(result[0]["role"], "assistant")
        self.assertEqual(result[0]["content"], "Checking files...")
        self.assertEqual(len(result[0]["tool_calls"]), 1)
        self.assertEqual(result[0]["tool_calls"][0]["id"], "call_123")
        self.assertEqual(result[0]["tool_calls"][0]["function"]["name"], "ls")
        self.assertEqual(json.loads(result[0]["tool_calls"][0]["function"]["arguments"]), {"path": "."})

        # 2: Tool result message
        self.assertEqual(result[1]["role"], "tool")
        self.assertEqual(result[1]["tool_call_id"], "call_123")
        self.assertEqual(result[1]["content"], "file1.txt\nfile2.txt")

        # 3: Remaining user text message
        self.assertEqual(result[2]["role"], "user")
        self.assertEqual(result[2]["content"], "Now delete file2.txt")

    def test_translate_openai_response_non_streaming(self):
        openai_resp = {
            "choices": [{
                "message": {
                    "role": "assistant",
                    "content": "Done!",
                    "tool_calls": [{
                        "id": "tc_1",
                        "type": "function",
                        "function": {
                            "name": "edit",
                            "arguments": "{\"file\": \"a.py\"}"
                        }
                    }]
                },
                "finish_reason": "tool_calls"
            }],
            "usage": {
                "prompt_tokens": 50,
                "completion_tokens": 20
            }
        }
        translated = translate_openai_response(openai_resp, "claude-sonnet-4-20250514")
        self.assertEqual(translated["role"], "assistant")
        self.assertEqual(translated["stop_reason"], "tool_use")
        self.assertEqual(len(translated["content"]), 2)
        self.assertEqual(translated["content"][0], {"type": "text", "text": "Done!"})
        self.assertEqual(translated["content"][1]["type"], "tool_use")
        self.assertEqual(translated["content"][1]["id"], "tc_1")
        self.assertEqual(translated["content"][1]["name"], "edit")
        self.assertEqual(translated["content"][1]["input"], {"file": "a.py"})
        self.assertEqual(translated["usage"]["input_tokens"], 50)
        self.assertEqual(translated["usage"]["output_tokens"], 20)

    def test_stream_translator_sse_sequence(self):
        """Test that StreamTranslator emits valid Anthropic SSE events in exact order."""
        buf = io.BytesIO()
        translator = StreamTranslator(buf, "claude-sonnet-4-20250514")

        # Chunk 1: text delta
        c1 = {"choices": [{"delta": {"content": "Hello"}}]}
        translator.process_chunk(c1)

        # Chunk 2: more text
        c2 = {"choices": [{"delta": {"content": " world!"}}]}
        translator.process_chunk(c2)

        # Chunk 3: finish with stop
        c3 = {
            "choices": [{"delta": {}, "finish_reason": "stop"}],
            "usage": {"prompt_tokens": 10, "completion_tokens": 5}
        }
        translator.process_chunk(c3)
        translator.finish()

        raw_output = buf.getvalue().decode("utf-8")
        events = []
        for block in raw_output.strip().split("\n\n"):
            lines = block.strip().split("\n")
            ev_type = lines[0].replace("event: ", "").strip()
            data = json.loads(lines[1].replace("data: ", "").strip())
            events.append((ev_type, data))

        event_names = [e[0] for e in events]
        expected_names = [
            "message_start",
            "ping",
            "content_block_start",
            "content_block_delta",
            "content_block_delta",
            "content_block_stop",
            "message_delta",
            "message_stop"
        ]
        self.assertEqual(event_names, expected_names)

        # Check content block text
        deltas = [e[1]["delta"]["text"] for e in events if e[0] == "content_block_delta"]
        self.assertEqual(deltas, ["Hello", " world!"])

        # Check usage in message_delta
        msg_delta = [e[1] for e in events if e[0] == "message_delta"][0]
        self.assertEqual(msg_delta["usage"]["output_tokens"], 5)
        self.assertEqual(msg_delta["delta"]["stop_reason"], "end_turn")

    def test_stream_translator_tool_call(self):
        """Test streaming tool call with multi-part arguments."""
        buf = io.BytesIO()
        translator = StreamTranslator(buf, "claude-sonnet-4-20250514")

        # Chunk 1: Tool call start
        c1 = {
            "choices": [{
                "delta": {
                    "tool_calls": [{
                        "index": 0,
                        "id": "call_abc",
                        "function": {"name": "calculator", "arguments": "{\"expr\":"}
                    }]
                }
            }]
        }
        translator.process_chunk(c1)

        # Chunk 2: Tool call argument delta
        c2 = {
            "choices": [{
                "delta": {
                    "tool_calls": [{
                        "index": 0,
                        "function": {"arguments": " \"2+2\"}"}
                    }]
                }
            }]
        }
        translator.process_chunk(c2)

        # Chunk 3: Finish reason tool_calls
        c3 = {
            "choices": [{"delta": {}, "finish_reason": "tool_calls"}],
            "usage": {"prompt_tokens": 15, "completion_tokens": 12}
        }
        translator.process_chunk(c3)
        translator.finish()

        raw_output = buf.getvalue().decode("utf-8")
        events = []
        for block in raw_output.strip().split("\n\n"):
            lines = block.strip().split("\n")
            ev_type = lines[0].replace("event: ", "").strip()
            data = json.loads(lines[1].replace("data: ", "").strip())
            events.append((ev_type, data))

        event_names = [e[0] for e in events]
        self.assertIn("message_start", event_names)
        self.assertIn("content_block_start", event_names)
        self.assertIn("content_block_delta", event_names)
        self.assertIn("content_block_stop", event_names)
        self.assertIn("message_delta", event_names)
        self.assertIn("message_stop", event_names)

        # Verify stop reason is tool_use
        msg_delta = [e[1] for e in events if e[0] == "message_delta"][0]
        self.assertEqual(msg_delta["delta"]["stop_reason"], "tool_use")

    def test_providers_and_keys_loading(self):
        providers = load_providers_json()
        self.assertIn("groq", providers)
        self.assertIn("Openrouter", providers)

        saved_keys = load_keys_from_excel()
        # Should load keys from C:\My Space\Data\keys.xlsx
        self.assertTrue(len(saved_keys) > 0)
        groq_key = find_key_for_provider("groq", saved_keys)
        self.assertTrue(len(groq_key) > 0)


if __name__ == "__main__":
    unittest.main()

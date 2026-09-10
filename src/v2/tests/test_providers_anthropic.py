import json

import pytest

import providers.anthropic as anthropic_provider
from utils.ai_errors import AIClientError


class FakeConfig:
    ai_model = "claude-test"
    ai_api_key = "fake-key"


SETTINGS = {"ai_endpoint": "https://example.test/v1/messages", "ai_api_version": "2099-01-01"}


class FakeHTTPResponse:
    def __init__(self, body: bytes):
        self._body = body

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_provider_name_and_required_settings():
    assert anthropic_provider.PROVIDER_NAME == "anthropic"
    assert anthropic_provider.REQUIRED_SETTINGS == ["ai_endpoint", "ai_api_version"]


def test_send_builds_correct_request_and_returns_extracted_text(monkeypatch):
    captured = {}

    def fake_urlopen(request, timeout):
        captured["url"] = request.full_url
        captured["headers"] = {k.lower(): v for k, v in request.headers.items()}
        captured["body"] = json.loads(request.data.decode("utf-8"))
        captured["timeout"] = timeout
        response_body = json.dumps({
            "content": [{"type": "text", "text": "hello "}, {"type": "text", "text": "world"}]
        }).encode("utf-8")
        return FakeHTTPResponse(response_body)

    monkeypatch.setattr(anthropic_provider.urllib.request, "urlopen", fake_urlopen)

    result = anthropic_provider.send("judge this", FakeConfig(), SETTINGS)

    assert result == "hello world"
    assert captured["url"] == SETTINGS["ai_endpoint"]
    assert captured["headers"]["x-api-key"] == "fake-key"
    assert captured["headers"]["anthropic-version"] == "2099-01-01"
    assert captured["body"]["model"] == "claude-test"
    assert captured["body"]["messages"] == [{"role": "user", "content": "judge this"}]


def test_send_raises_ai_client_error_on_connection_failure(monkeypatch):
    def fake_urlopen(request, timeout):
        raise OSError("connection refused")

    monkeypatch.setattr(anthropic_provider.urllib.request, "urlopen", fake_urlopen)

    with pytest.raises(AIClientError):
        anthropic_provider.send("prompt", FakeConfig(), SETTINGS)


def test_send_raises_ai_client_error_on_malformed_json_response(monkeypatch):
    def fake_urlopen(request, timeout):
        return FakeHTTPResponse(b"not json")

    monkeypatch.setattr(anthropic_provider.urllib.request, "urlopen", fake_urlopen)

    with pytest.raises(AIClientError):
        anthropic_provider.send("prompt", FakeConfig(), SETTINGS)


def test_extract_text_concatenates_only_text_blocks():
    data = {"content": [
        {"type": "text", "text": "a"},
        {"type": "other", "text": "b"},
        {"type": "text", "text": "c"},
    ]}
    assert anthropic_provider._extract_text(data) == "ac"


def test_extract_text_raises_on_unexpected_shape():
    with pytest.raises(AIClientError):
        anthropic_provider._extract_text({"unexpected": "shape"})

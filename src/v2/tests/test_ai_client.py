import pytest

import utils.ai_client as ai_client
from utils.ai_client import AIClientError, send_batch, send_prompt
from utils.ai_client import test_connection as check_ai_connection


class FakeConfig:
    ai_model = "fake-model"
    ai_api_key = "fake-key"


# ---------------------------------------------------------------------------
# _dispatch — the seam between ai_client.py and the Provider Registry
# (FS section 2.6.3). Everything above this line mocks _dispatch itself to
# test send_prompt/send_batch/test_connection's retry/parsing/error-handling
# in isolation; these tests exercise _dispatch itself, with a fake provider
# module/registry standing in for a real one.
# ---------------------------------------------------------------------------

class FakeProviderModule:
    def __init__(self, response="OK"):
        self.response = response
        self.calls = []

    def send(self, prompt, config, settings):
        self.calls.append((prompt, config, settings))
        return self.response


class FakeProviderRegistry:
    def __init__(self, providers=None):
        self._providers = providers or {}

    def get(self, name):
        return self._providers.get(str(name).strip().lower())


def test_dispatch_calls_selected_providers_send_with_resolved_settings():
    provider = FakeProviderModule(response="hello")
    registry = FakeProviderRegistry({"anthropic": provider})

    class Cfg:
        ai_model = "m"
        ai_api_key = "k"
        ai_provider = "anthropic"
        ai_provider_settings = {"anthropic": {"ai_endpoint": "https://x", "ai_api_version": "v1"}}
        provider_registry = registry

    result = ai_client._dispatch(Cfg(), "hi")

    assert result == "hello"
    assert len(provider.calls) == 1
    prompt, config, settings = provider.calls[0]
    assert prompt == "hi"
    assert settings == {"ai_endpoint": "https://x", "ai_api_version": "v1"}


def test_dispatch_raises_when_no_provider_registry_on_config():
    class Cfg:
        ai_provider = "anthropic"
        ai_provider_settings = {}
        provider_registry = None

    with pytest.raises(AIClientError, match="provider registry"):
        ai_client._dispatch(Cfg(), "hi")


def test_dispatch_raises_when_provider_not_registered():
    registry = FakeProviderRegistry({})

    class Cfg:
        ai_provider = "not-a-real-provider"
        ai_provider_settings = {}
        provider_registry = registry

    with pytest.raises(AIClientError, match="not registered"):
        ai_client._dispatch(Cfg(), "hi")


def test_dispatch_provider_lookup_is_case_insensitive():
    provider = FakeProviderModule(response="hi back")
    registry = FakeProviderRegistry({"anthropic": provider})

    class Cfg:
        ai_model = "m"
        ai_api_key = "k"
        ai_provider = "Anthropic"
        ai_provider_settings = {"anthropic": {"ai_endpoint": "https://x", "ai_api_version": "v1"}}
        provider_registry = registry

    assert ai_client._dispatch(Cfg(), "hi") == "hi back"


def test_send_batch_parses_verdicts_and_comments(monkeypatch):
    def fake_call(config, prompt):
        return "P,F,P\n2: Item name doesn't match category"
    monkeypatch.setattr(ai_client, "_dispatch", fake_call)

    results = send_batch(
        FakeConfig(), "judge these", [{"a": 1}, {"a": 2}, {"a": 3}]
    )
    assert results == [(True, ""), (False, "Item name doesn't match category"), (True, "")]


def test_send_batch_empty_payloads_returns_empty_without_calling(monkeypatch):
    called = []
    monkeypatch.setattr(ai_client, "_dispatch", lambda c, p: called.append(1) or "P")
    assert send_batch(FakeConfig(), "prompt", []) == []
    assert called == []


def test_send_batch_includes_reference_context_in_prompt(monkeypatch):
    captured = {}

    def fake_call(config, prompt):
        captured["prompt"] = prompt
        return "P"

    monkeypatch.setattr(ai_client, "_dispatch", fake_call)
    send_batch(FakeConfig(), "judge", [{"a": 1}], reference_context=["Widgets", "Fasteners"])
    assert "Widgets" in captured["prompt"]
    assert "Fasteners" in captured["prompt"]


def test_send_batch_verdict_count_mismatch_raises(monkeypatch):
    monkeypatch.setattr(ai_client, "_dispatch", lambda c, p: "P,P")
    with pytest.raises(AIClientError):
        send_batch(FakeConfig(), "prompt", [{"a": 1}, {"a": 2}, {"a": 3}])


def test_send_batch_invalid_verdict_character_raises(monkeypatch):
    monkeypatch.setattr(ai_client, "_dispatch", lambda c, p: "P,X")
    with pytest.raises(AIClientError):
        send_batch(FakeConfig(), "prompt", [{"a": 1}, {"a": 2}])


def test_send_batch_empty_response_raises(monkeypatch):
    monkeypatch.setattr(ai_client, "_dispatch", lambda c, p: "")
    with pytest.raises(AIClientError):
        send_batch(FakeConfig(), "prompt", [{"a": 1}])


def test_send_prompt_retries_then_succeeds(monkeypatch):
    calls = {"n": 0}

    def flaky(config, prompt):
        calls["n"] += 1
        if calls["n"] < 3:
            raise RuntimeError("transient network error")
        return "OK"

    monkeypatch.setattr(ai_client, "_dispatch", flaky)
    result = send_prompt(FakeConfig(), "ping", max_retries=2, retry_delay=0)
    assert result == "OK"
    assert calls["n"] == 3


def test_send_prompt_raises_ai_client_error_after_exhausting_retries(monkeypatch):
    def always_fails(config, prompt):
        raise RuntimeError("down")

    monkeypatch.setattr(ai_client, "_dispatch", always_fails)
    with pytest.raises(AIClientError):
        send_prompt(FakeConfig(), "ping", max_retries=1, retry_delay=0)


def test_send_batch_propagates_ai_client_error_from_call(monkeypatch):
    def always_fails(config, prompt):
        raise RuntimeError("down")

    monkeypatch.setattr(ai_client, "_dispatch", always_fails)
    with pytest.raises(AIClientError):
        send_batch(FakeConfig(), "prompt", [{"a": 1}], reference_context=None)


def test_ai_connectivity_check_uses_same_request_path(monkeypatch):
    captured = {}

    def fake_call(config, prompt):
        captured["prompt"] = prompt
        return "OK"

    monkeypatch.setattr(ai_client, "_dispatch", fake_call)
    check_ai_connection(FakeConfig(), "Respond with OK.")
    assert captured["prompt"] == "Respond with OK."


def test_ai_connectivity_check_raises_on_failure(monkeypatch):
    def always_fails(config, prompt):
        raise RuntimeError("bad key")

    monkeypatch.setattr(ai_client, "_dispatch", always_fails)
    with pytest.raises(AIClientError):
        check_ai_connection(FakeConfig(), "Respond with OK.")

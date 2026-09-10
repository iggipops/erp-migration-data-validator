"""
Anthropic provider module — Provider Registry (FS section 2.6.3), reference
implementation (Appendix F). This is the only place in the framework that
knows the Anthropic Messages API's specific request/response shape.

Registered automatically by registry/provider_registry.py; never imported
directly by callers. utils/ai_client.py reaches this module only through
config.provider_registry.get(config.ai_provider), never by name.
"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

from utils.ai_errors import AIClientError

PROVIDER_NAME = "anthropic"

# Both mandatory in ai_provider_settings.anthropic (config.yaml, section 3.4)
# whenever ai_provider: anthropic is selected and ai_enabled = true — checked
# automatically at Step 2 (8.2), the same pattern REQUIRED_PARAMS gives
# custom functions (section 7.2.1).
REQUIRED_SETTINGS = ["ai_endpoint", "ai_api_version"]

DEFAULT_MAX_TOKENS = 1024
DEFAULT_TIMEOUT_SECONDS = 60


def send(prompt: str, config, settings: dict) -> str:
    """
    Sends `prompt` to the Anthropic Messages API and returns the raw text
    response. `config` supplies the fields every provider needs in common
    (ai_model, ai_api_key); `settings` is this provider's own resolved
    ai_provider_settings.anthropic block (ai_endpoint, ai_api_version),
    already checked against REQUIRED_SETTINGS by the config loader.
    """
    request_body = json.dumps({
        "model": config.ai_model,
        "max_tokens": DEFAULT_MAX_TOKENS,
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")

    request = urllib.request.Request(
        settings["ai_endpoint"],
        data=request_body,
        headers={
            "x-api-key": config.ai_api_key,
            "anthropic-version": settings["ai_api_version"],
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=DEFAULT_TIMEOUT_SECONDS) as response:
            data = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, OSError, ValueError) as exc:
        raise AIClientError(f"AI service request failed: {exc}") from exc

    return _extract_text(data)


def _extract_text(data: dict) -> str:
    try:
        return "".join(
            block["text"] for block in data["content"] if block.get("type") == "text"
        )
    except (KeyError, IndexError, TypeError) as exc:
        raise AIClientError(f"Unexpected AI response shape: {data}") from exc

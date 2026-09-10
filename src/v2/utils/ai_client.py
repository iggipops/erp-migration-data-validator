"""
Shared AI client module (FS section 6.8).

Provides the single request path used both by real AI-flavored validation
batches (via send_batch) and by the ai_connection_test_prompt pre-flight
check at Step 2 (via test_connection, FS sections 3.4/3.8/8.2.1). Individual
AI-flavored custom functions (e.g. functions/custom/ai_itemcategoryfit.py)
only decide what data to gather and what prompt to send — building the
request, calling the configured AI service, parsing the standardized
per-row verdict response, and retry/error handling all live here.

This module is itself provider-agnostic: it dispatches to whichever
provider module config.ai_provider selects, looked up in the Provider
Registry (config.provider_registry, section 2.6.3) — every provider
module is the only place that knows its own HTTP/SDK specifics. Neither
this module nor any caller changes based on which provider is active.
"""
from __future__ import annotations

import json
import time

from utils.ai_errors import AIClientError, AIRulePartialError
from utils.logger import get_logger

__all__ = [
    "AIClientError",
    "AIRulePartialError",
    "send_prompt",
    "send_batch",
    "test_connection",
]

DEFAULT_MAX_RETRIES = 2
DEFAULT_RETRY_DELAY_SECONDS = 1.0


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def send_prompt(
    config,
    prompt: str,
    max_retries: int = DEFAULT_MAX_RETRIES,
    retry_delay: float = DEFAULT_RETRY_DELAY_SECONDS,
) -> str:
    """
    Send `prompt` to the configured AI provider (config.ai_provider, via
    config.provider_registry) and return the raw text response. Retries
    transient failures up to `max_retries` times before raising
    AIClientError.
    """
    logger = get_logger()
    last_exc: Exception | None = None

    for attempt in range(1, max_retries + 2):
        try:
            return _dispatch(config, prompt)
        except Exception as exc:
            last_exc = exc
            logger.warning(f"AI request attempt {attempt} failed: {exc}")
            if attempt <= max_retries:
                time.sleep(retry_delay)

    raise AIClientError(
        f"AI request failed after {max_retries + 1} attempt(s): {last_exc}"
    ) from last_exc


def send_batch(
    config,
    prompt: str,
    row_payloads: list[dict],
    reference_context: list[str] | None = None,
) -> list[tuple[bool, str]]:
    """
    Send one batch of rows to the AI service and parse the standardized
    per-row verdict response (FS section 6.8): a compact ordered list
    such as "P,P,F,P,F", with comment text present only for failing rows.

    `row_payloads` is one JSON-serializable dict per row, in order.
    `reference_context`, if given, is sent once for the whole batch
    (e.g. a full reference category list) rather than once per row.

    Returns a list of (passed, comment) tuples, one per row_payloads
    entry, in the same order. Raises AIClientError if the call fails or
    the response can't be parsed into exactly len(row_payloads) verdicts.
    """
    if not row_payloads:
        return []

    request_prompt = _build_batch_prompt(prompt, row_payloads, reference_context)
    raw_response = send_prompt(config, request_prompt)
    return _parse_batch_response(raw_response, len(row_payloads))


def test_connection(config, test_prompt: str) -> None:
    """
    Pre-flight connectivity check (FS sections 3.4/3.8, Step 2/8.2.1).
    Sends `test_prompt` through the same request path as a real batch,
    dispatched via the Provider Registry like any other AI call — no
    function-specific code is involved. Raises AIClientError on failure.
    """
    send_prompt(config, test_prompt)


# ---------------------------------------------------------------------------
# Dispatch — the only place that touches the Provider Registry
# ---------------------------------------------------------------------------

def _dispatch(config, prompt: str) -> str:
    provider_registry = getattr(config, "provider_registry", None)
    if provider_registry is None:
        raise AIClientError(
            "No provider registry available on config — the framework "
            "must build the Provider Registry (Step 1, FS 8.1) and attach "
            "it to config before any AI call can be dispatched"
        )

    provider = provider_registry.get(config.ai_provider)
    if provider is None:
        raise AIClientError(
            f"AI provider '{config.ai_provider}' is not registered"
        )

    settings = (config.ai_provider_settings or {}).get(
        str(config.ai_provider).strip().lower(), {}
    )
    return provider.send(prompt, config, settings)


# ---------------------------------------------------------------------------
# Batch prompt construction / response parsing — provider-agnostic
# ---------------------------------------------------------------------------

def _build_batch_prompt(
    prompt: str,
    row_payloads: list[dict],
    reference_context: list[str] | None,
) -> str:
    lines = [
        prompt,
        "",
        "Respond with exactly one verdict per row, in order, as a single "
        "comma-separated line of P (pass) or F (fail) - e.g. P,P,F,P. "
        "After the verdict line, for each F row only, add one line "
        "'N: <comment>' where N is the row's 1-based position in this "
        "batch and <comment> is your explanation.",
        "",
        "Rows:",
    ]
    for i, payload in enumerate(row_payloads, start=1):
        lines.append(f"{i}. {json.dumps(payload, default=str)}")

    if reference_context:
        lines.append("")
        lines.append("Reference list:")
        for item in reference_context:
            lines.append(f"- {item}")

    return "\n".join(lines)


def _parse_batch_response(raw_response: str, expected_count: int) -> list[tuple[bool, str]]:
    lines = [ln.strip() for ln in (raw_response or "").strip().splitlines() if ln.strip()]
    if not lines:
        raise AIClientError("AI response was empty")

    verdict_line = lines[0]
    verdicts = [v.strip().upper() for v in verdict_line.split(",") if v.strip()]

    if len(verdicts) != expected_count:
        raise AIClientError(
            f"AI response verdict count ({len(verdicts)}) does not match "
            f"expected row count ({expected_count}): {verdict_line!r}"
        )

    if any(v not in ("P", "F") for v in verdicts):
        raise AIClientError(
            f"AI response contains invalid verdict character(s): {verdict_line!r}"
        )

    comments: dict[int, str] = {}
    for line in lines[1:]:
        if ":" not in line:
            continue
        idx_str, comment = line.split(":", 1)
        idx_str = idx_str.strip()
        if idx_str.isdigit():
            comments[int(idx_str)] = comment.strip()

    return [
        (verdict == "P", comments.get(i, "") if verdict == "F" else "")
        for i, verdict in enumerate(verdicts, start=1)
    ]

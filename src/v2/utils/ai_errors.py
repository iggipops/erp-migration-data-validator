"""
AI-related exception types, kept separate from utils/ai_client.py so that
provider modules (providers/*.py) can raise them without creating a circular
import: ai_client.py dispatches to provider modules via the Provider
Registry, so provider modules must not import back from ai_client.py.
"""
from __future__ import annotations


class AIClientError(Exception):
    """
    Raised when a request to the AI service fails (after retries) or its
    response cannot be parsed into the expected verdict format. Callers
    catch this and translate it into a runtime failure scoped to just the
    batch or check in question (FS section 8.8.2) — never a crash of the
    whole run.
    """


class AIRulePartialError(Exception):
    """
    Raised by an AI ValidationType dispatcher (validators/builtin/ai.py)
    when at least one batch succeeded and at least one batch failed for
    the same rule (FS section 6.8). Issues for successful-batch rows have
    already been recorded by the time this is raised — it exists only to
    signal the rule's overall status should be 'partial', not 'failed'.

    `rows_unchecked` is set by the raiser (the number of rows whose batch
    failed and therefore have no verdict). `issue_count` is filled in by
    validation_engine.py._dispatch(), which is the only place that knows
    the before/after issue counter delta.
    """
    rows_unchecked: int = 0
    issue_count: int = 0

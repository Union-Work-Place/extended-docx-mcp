"""Common response envelope helpers for FastMCP tools."""

from __future__ import annotations

import datetime as dt
from functools import wraps
from typing import Any, Callable


METADATA_FIELDS = {
    "path",
    "saved_to",
    "engine",
    "track_changes_requested",
    "track_changes_supported",
    "created_from_missing",
}


def _timestamp() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def _normalize_success_payload(payload: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    if isinstance(payload, dict):
        result = dict(payload)
        metadata = {key: result.pop(key) for key in tuple(result) if key in METADATA_FIELDS}
        return result, metadata
    return {"value": payload}, {}


def success_response(operation: str, payload: Any) -> dict[str, Any]:
    """Wrap a tool result in the shared success envelope."""

    result, metadata = _normalize_success_payload(payload)
    metadata["timestamp"] = _timestamp()
    return {
        "status": "ok",
        "operation": operation,
        "result": result,
        "metadata": metadata,
    }


def error_response(operation: str, exc: Exception) -> dict[str, Any]:
    """Convert an exception into the shared error envelope."""

    if isinstance(exc, FileNotFoundError):
        error_code = "document_not_found"
        suggestion = "Check the DOCX path and the EXTENDED_DOCX_MCP_DEFAULT_DIR setting."
    elif isinstance(exc, IndexError):
        error_code = "index_out_of_range"
        suggestion = "Check the index and inspect the document structure first with read_docx or a list_* tool."
    elif isinstance(exc, ValueError):
        error_code = "invalid_input"
        suggestion = "Review the tool arguments and retry with valid values."
    else:
        error_code = "internal_error"
        suggestion = "Retry the request or inspect the document for unsupported OOXML structures."
    return {
        "status": "error",
        "operation": operation,
        "error": {
            "code": error_code,
            "message": str(exc),
            "suggestion": suggestion,
            "type": exc.__class__.__name__,
        },
        "metadata": {
            "timestamp": _timestamp(),
        },
    }


def tool_response(operation: str) -> Callable[[Callable[..., Any]], Callable[..., dict[str, Any]]]:
    """Wrap a FastMCP tool with the shared success/error envelope."""

    def decorator(func: Callable[..., Any]) -> Callable[..., dict[str, Any]]:
        @wraps(func)
        def wrapped(*args: Any, **kwargs: Any) -> dict[str, Any]:
            try:
                return success_response(operation, func(*args, **kwargs))
            except (FileNotFoundError, IndexError, RuntimeError, ValueError) as exc:  # pragma: no cover - integration-tested
                return error_response(operation, exc)

        return wrapped

    return decorator

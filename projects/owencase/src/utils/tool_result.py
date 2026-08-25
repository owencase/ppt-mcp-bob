"""Shared execution and JSON serialization for MCP tool handlers."""

from __future__ import annotations

import json
from collections.abc import Callable
from typing import Any

from utils.com_wrapper import ppt


def execute_json(
    error_prefix: str | None,
    operation: Callable[..., Any],
    *args: Any,
) -> str:
    """Run a COM operation and return a consistent JSON result."""
    try:
        return json.dumps(ppt.execute(operation, *args))
    except Exception as exc:
        message = f"{error_prefix}{exc}" if error_prefix else str(exc)
        return json.dumps({"error": message})

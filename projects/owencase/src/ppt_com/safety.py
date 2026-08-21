"""Session-level safety policy for IBM Bob PowerPoint operations."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from threading import RLock
from typing import Optional


@dataclass
class WorkMode:
    """Mutable policy applied to every tool call in the current MCP session."""

    allow_create: bool = False
    require_preconditions: bool = True
    expected_presentation_full_name: Optional[str] = None
    expected_slide_count: Optional[int] = None


_mode = WorkMode()
_lock = RLock()


def get_work_mode() -> WorkMode:
    """Return a copy so callers cannot mutate policy without validation."""
    with _lock:
        return WorkMode(**asdict(_mode))


def set_work_mode(
    *,
    allow_create: bool,
    require_preconditions: bool,
    expected_presentation_full_name: Optional[str] = None,
    expected_slide_count: Optional[int] = None,
) -> WorkMode:
    """Replace the current policy and return the resulting mode."""
    global _mode
    with _lock:
        _mode = WorkMode(
            allow_create=allow_create,
            require_preconditions=require_preconditions,
            expected_presentation_full_name=expected_presentation_full_name,
            expected_slide_count=expected_slide_count,
        )
        return WorkMode(**asdict(_mode))


def work_mode_dict() -> dict:
    return asdict(get_work_mode())

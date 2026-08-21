"""IBM Bob-facing guards shared by every registered MCP tool."""

from __future__ import annotations

import inspect
import json
import ntpath
import os
from functools import wraps
from typing import Any, Awaitable, Callable, get_type_hints

from pydantic import BaseModel, ConfigDict

from ppt_com.safety import get_work_mode


CREATE_TOOL_NAMES = {
    "ppt_add_animation",
    "ppt_add_audio",
    "ppt_add_chart",
    "ppt_add_comment",
    "ppt_add_connector",
    "ppt_add_hyperlink",
    "ppt_add_line",
    "ppt_add_picture",
    "ppt_add_picture_from_url",
    "ppt_add_section",
    "ppt_add_shape",
    "ppt_add_slide",
    "ppt_add_smartart",
    "ppt_add_svg_icon",
    "ppt_add_table",
    "ppt_add_table_column",
    "ppt_add_table_row",
    "ppt_add_textbox",
    "ppt_add_video",
    "ppt_copy_shape_to_slide",
    "ppt_copy_slide",
    "ppt_create_presentation",
    "ppt_duplicate_shape",
    "ppt_duplicate_slide",
    "ppt_insert_node",
}

PRECONDITION_EXEMPT_TOOLS = {
    "ppt_activate_presentation",
    "ppt_connect",
    "ppt_create_presentation",
    "ppt_get_work_mode",
    "ppt_list_presentations",
    "ppt_open_presentation",
    "ppt_set_work_mode",
    "ppt_set_window_state",
}

# These legacy tools mutate shape geometry or identity without checking the
# observed shape state. They remain available when preconditions are explicitly
# disabled, but safe IBM Bob editing must use the guarded high-level operations.
UNPRECONDITIONED_SHAPE_TOOLS = {
    "ppt_align_shapes",
    "ppt_delete_shape",
    "ppt_distribute_shapes",
    "ppt_set_shape_zorder",
    "ppt_update_shape",
}

OUTPUT_PATH_FIELDS = {
    "ppt_export_images": "output_dir",
    "ppt_export_pdf": "file_path",
    "ppt_export_shape": "file_path",
    "ppt_save_presentation_as": "file_path",
}


class BobToolError(RuntimeError):
    """An MCP-visible failure with explicit retry guidance."""

    def __init__(self, error: str, *, retryable: bool, hint: str):
        self.payload = {
            "error": error,
            "retryable": retryable,
            "hint": hint,
        }
        super().__init__(json.dumps(self.payload, ensure_ascii=False))


def _normalise_windows_path(value: str) -> str:
    return ntpath.normcase(ntpath.normpath(value.strip()))


def _target_state_impl() -> dict[str, Any]:
    from utils.com_wrapper import ppt

    pres = ppt._get_pres_impl()
    app = ppt._get_app_impl()
    presentation_index = None
    for index in range(1, app.Presentations.Count + 1):
        if _normalise_windows_path(str(app.Presentations(index).FullName)) == (
            _normalise_windows_path(str(pres.FullName))
        ):
            presentation_index = index
            break
    return {
        "full_name": str(pres.FullName),
        "name": str(pres.Name),
        "presentation_index": presentation_index,
        "slide_count": int(pres.Slides.Count),
    }


def validate_session_preconditions(
    tool_name: str,
    *,
    mutating: bool,
    params: Any = None,
) -> None:
    """Block unsafe calls before they reach COM."""
    mode = get_work_mode()

    if tool_name in CREATE_TOOL_NAMES and not mode.allow_create:
        raise BobToolError(
            f"Creation is disabled for '{tool_name}'.",
            retryable=False,
            hint=(
                "For move, resize, edit, or replacement requests, modify existing "
                "shape IDs instead of calling ppt_add_*, ppt_copy_*, or ppt_duplicate_*. "
                "Only call ppt_set_work_mode with allow_create=true when the user "
                "explicitly requested new content."
            ),
        )

    if not mutating or tool_name in PRECONDITION_EXEMPT_TOOLS:
        return
    if not mode.require_preconditions:
        return
    if not mode.expected_presentation_full_name:
        raise BobToolError(
            f"'{tool_name}' requires an expected presentation before editing.",
            retryable=False,
            hint=(
                "Call ppt_activate_presentation, inspect it, then call ppt_set_work_mode "
                "with expected_presentation_full_name and expected_slide_count."
            ),
        )

    from utils.com_wrapper import ppt

    state = ppt.execute(_target_state_impl)
    expected = _normalise_windows_path(mode.expected_presentation_full_name)
    actual = _normalise_windows_path(state["full_name"])
    if actual != expected:
        raise BobToolError(
            f"Presentation precondition failed: expected '{mode.expected_presentation_full_name}', "
            f"but the active target is '{state['full_name']}'.",
            retryable=False,
            hint="Activate the intended presentation and set work mode again; do not edit another open deck.",
        )
    if (
        mode.expected_slide_count is not None
        and state["slide_count"] != mode.expected_slide_count
    ):
        raise BobToolError(
            f"Slide-count precondition failed: expected {mode.expected_slide_count}, "
            f"found {state['slide_count']}.",
            retryable=False,
            hint="Inspect the presentation again and confirm whether its structure changed before retrying.",
        )

    selector_index = getattr(params, "presentation_index", None)
    selector_name = getattr(params, "presentation_name", None)
    if selector_index is not None and selector_index != state["presentation_index"]:
        raise BobToolError(
            f"Presentation selector {selector_index} does not match the locked target.",
            retryable=False,
            hint="Remove the selector or use the locked presentation's current index.",
        )
    if selector_name is not None and selector_name.lower() not in {
        state["name"].lower(),
        state["full_name"].lower(),
    }:
        raise BobToolError(
            f"Presentation selector '{selector_name}' does not match the locked target.",
            retryable=False,
            hint="Remove the selector or use the locked presentation's exact name or full path.",
        )

    destination_index = getattr(params, "to_presentation_index", None)
    destination_name = getattr(params, "to_presentation_name", None)
    if (
        destination_index is not None
        and destination_index != state["presentation_index"]
    ):
        raise BobToolError(
            "The copy destination does not match the locked target presentation.",
            retryable=False,
            hint="Set the destination to the locked target or activate the intended destination first.",
        )
    if destination_name is not None and destination_name.lower() not in {
        state["name"].lower(),
        state["full_name"].lower(),
    }:
        raise BobToolError(
            "The copy destination does not match the locked target presentation.",
            retryable=False,
            hint="Set the destination to the locked target or activate the intended destination first.",
        )

    if tool_name in UNPRECONDITIONED_SHAPE_TOOLS:
        raise BobToolError(
            f"'{tool_name}' does not accept shape-state preconditions.",
            retryable=False,
            hint=(
                "Use ppt_transform_shapes for move, resize, or rotation, "
                "ppt_delete_shapes for deletion, and ppt_replace_visual for an "
                "authentic existing replacement. These tools require shape_id "
                "and observed shape properties."
            ),
        )


def validate_output_path(tool_name: str, args: tuple[Any, ...], kwargs: dict[str, Any]) -> None:
    field_name = OUTPUT_PATH_FIELDS.get(tool_name)
    if not field_name:
        return
    params = args[0] if args else kwargs.get("params")
    if params is None:
        return
    value = getattr(params, field_name, None)
    if not value:
        return
    if not ntpath.isabs(value):
        raise BobToolError(
            f"{tool_name} requires an absolute {field_name}.",
            retryable=False,
            hint="Use an absolute path inside the configured output directory.",
        )

    mode = get_work_mode()
    allowed_root = os.getenv("PPT_MCP_OUTPUT_DIR")
    if not allowed_root and mode.expected_presentation_full_name:
        expected = mode.expected_presentation_full_name
        if not expected.lower().startswith(("http://", "https://")):
            allowed_root = ntpath.dirname(expected)
    if not allowed_root:
        raise BobToolError(
            f"No output trust boundary is configured for '{tool_name}'.",
            retryable=False,
            hint=(
                "Set PPT_MCP_OUTPUT_DIR in the MCP environment, or set work mode "
                "with a local expected_presentation_full_name."
            ),
        )

    candidate = _normalise_windows_path(value)
    root = _normalise_windows_path(allowed_root)
    try:
        inside = ntpath.commonpath([candidate, root]) == root
    except ValueError:
        inside = False
    if not inside:
        raise BobToolError(
            f"Output path '{value}' is outside the allowed directory '{allowed_root}'.",
            retryable=False,
            hint="Choose a path inside the allowed output directory.",
        )


def _classify_exception(exc: Exception) -> tuple[bool, str]:
    message = str(exc).lower()
    if any(token in message for token in ("busy", "call rejected", "retry later", "temporarily")):
        return True, "PowerPoint is temporarily busy. Close modal dialogs and retry once."
    if isinstance(exc, ConnectionError):
        return False, "Open PowerPoint, open the target deck, then call ppt_connect and ppt_activate_presentation."
    if isinstance(exc, (ValueError, FileNotFoundError, PermissionError)):
        return False, "Correct the input or permissions before retrying; repeating the same call will not help."
    return False, "Inspect the presentation and inputs before retrying; do not repeat the same mutating call blindly."


def promote_error_result(tool_name: str, result: Any) -> Any:
    """Turn legacy JSON error strings into real MCP tool failures."""
    if not isinstance(result, str):
        return result
    try:
        payload = json.loads(result)
    except (TypeError, json.JSONDecodeError):
        return result
    if not isinstance(payload, dict) or "error" not in payload:
        return result
    retryable = bool(payload.get("retryable", False))
    hint = str(
        payload.get("hint")
        or "Correct the input or inspect the target state before retrying."
    )
    raise BobToolError(
        f"{tool_name}: {payload['error']}",
        retryable=retryable,
        hint=hint,
    )


def make_strict_input_models(func: Callable[..., Any]) -> None:
    """Apply extra='forbid' before FastMCP builds each tool schema."""
    try:
        hints = get_type_hints(func)
    except Exception:
        hints = {}
    for parameter in inspect.signature(func).parameters.values():
        annotation = hints.get(parameter.name, parameter.annotation)
        if not inspect.isclass(annotation) or not issubclass(annotation, BaseModel):
            continue
        annotation.model_config = ConfigDict(
            **{**annotation.model_config, "extra": "forbid"}
        )
        annotation.model_rebuild(force=True)


def guarded_tool_decorator(
    register: Callable[[Callable[..., Any]], Any],
    *,
    tool_name: str,
    annotations: dict[str, Any] | None,
) -> Callable[[Callable[..., Awaitable[Any]]], Any]:
    """Wrap a FastMCP decorator with Bob policy and error promotion."""
    mutating = not bool((annotations or {}).get("readOnlyHint", False))

    def decorator(func: Callable[..., Awaitable[Any]]) -> Any:
        make_strict_input_models(func)

        @wraps(func)
        async def guarded(*args: Any, **kwargs: Any) -> Any:
            try:
                params = args[0] if args else kwargs.get("params")
                validate_session_preconditions(
                    tool_name,
                    mutating=mutating,
                    params=params,
                )
                validate_output_path(tool_name, args, kwargs)
                result = await func(*args, **kwargs)
                return promote_error_result(tool_name, result)
            except BobToolError:
                raise
            except Exception as exc:
                retryable, hint = _classify_exception(exc)
                raise BobToolError(
                    f"{tool_name}: {exc}",
                    retryable=retryable,
                    hint=hint,
                ) from exc

        return register(guarded)

    return decorator

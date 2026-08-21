"""High-level, preconditioned PowerPoint operations for IBM Bob."""

from __future__ import annotations

import json
import ntpath
import uuid
from collections import OrderedDict
from threading import RLock
from typing import Literal, Optional

from pydantic import BaseModel, ConfigDict, Field, model_validator

from ppt_com.safety import set_work_mode, work_mode_dict
from utils.com_wrapper import ppt


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class SetWorkModeInput(StrictModel):
    allow_create: bool = False
    require_preconditions: bool = True
    expected_presentation_full_name: Optional[str] = None
    expected_slide_count: Optional[int] = Field(default=None, ge=0)

    @model_validator(mode="after")
    def require_target_for_safe_editing(self):
        if self.require_preconditions and not self.expected_presentation_full_name:
            raise ValueError(
                "expected_presentation_full_name is required when "
                "require_preconditions=true"
            )
        return self


class ShapePrecondition(StrictModel):
    expected_name: Optional[str] = None
    expected_left: Optional[float] = None
    expected_top: Optional[float] = None
    expected_width: Optional[float] = Field(default=None, gt=0)
    expected_height: Optional[float] = Field(default=None, gt=0)
    expected_text: Optional[str] = None
    tolerance: float = Field(default=0.5, ge=0, le=10)

    @model_validator(mode="after")
    def require_observed_state(self):
        observed = (
            self.expected_name,
            self.expected_left,
            self.expected_top,
            self.expected_width,
            self.expected_height,
            self.expected_text,
        )
        if all(value is None for value in observed):
            raise ValueError("At least one expected shape property is required")
        return self


class ShapeTransform(StrictModel):
    shape_id: int = Field(..., ge=1)
    precondition: ShapePrecondition
    left: Optional[float] = None
    top: Optional[float] = None
    width: Optional[float] = Field(default=None, gt=0)
    height: Optional[float] = Field(default=None, gt=0)
    rotation: Optional[float] = Field(default=None, ge=0, le=360)

    @model_validator(mode="after")
    def require_change(self):
        if all(
            value is None
            for value in (self.left, self.top, self.width, self.height, self.rotation)
        ):
            raise ValueError("At least one transform property is required")
        return self


class TransformShapesInput(StrictModel):
    slide_index: int = Field(..., ge=1)
    transforms: list[ShapeTransform] = Field(..., min_length=1)

    @model_validator(mode="after")
    def unique_shape_ids(self):
        ids = [item.shape_id for item in self.transforms]
        if len(ids) != len(set(ids)):
            raise ValueError("Each shape_id may appear only once")
        return self


class DeleteShapesInput(StrictModel):
    slide_index: int = Field(..., ge=1)
    targets: dict[int, ShapePrecondition] = Field(..., min_length=1)

    @model_validator(mode="after")
    def validate_shape_ids(self):
        invalid = sorted(shape_id for shape_id in self.targets if shape_id < 1)
        if invalid:
            raise ValueError(f"shape_id values must be positive: {invalid}")
        return self


class ReplaceVisualInput(StrictModel):
    slide_index: int = Field(..., ge=1)
    target_shape_ids: list[int] = Field(..., min_length=1)
    target_preconditions: dict[int, ShapePrecondition]
    replacement_shape_id: int = Field(..., ge=1)
    replacement_precondition: ShapePrecondition
    fit_mode: Literal["contain", "fill", "keep_size"] = "contain"
    delete_targets: bool = True

    @model_validator(mode="after")
    def validate_targets(self):
        if len(self.target_shape_ids) != len(set(self.target_shape_ids)):
            raise ValueError("target_shape_ids must be unique")
        if self.replacement_shape_id in self.target_shape_ids:
            raise ValueError("replacement_shape_id cannot also be a target")
        missing = set(self.target_shape_ids) - set(self.target_preconditions)
        extra = set(self.target_preconditions) - set(self.target_shape_ids)
        if missing or extra:
            raise ValueError(
                f"target_preconditions must match target_shape_ids; "
                f"missing={sorted(missing)}, extra={sorted(extra)}"
            )
        return self


class SnapshotInput(StrictModel):
    slide_indices: Optional[list[int]] = None


class DiffSnapshotInput(StrictModel):
    snapshot_id: str = Field(..., min_length=1)


class ValidatePresentationInput(StrictModel):
    slide_indices: Optional[list[int]] = None
    min_font_size: float = Field(default=16.0, gt=0)
    minimum_frame_size: float = Field(default=4.0, gt=0)


_snapshots: OrderedDict[str, dict] = OrderedDict()
_snapshot_lock = RLock()
_MAX_SNAPSHOTS = 20


def _normalise_windows_path(value: str) -> str:
    return ntpath.normcase(ntpath.normpath(value.strip()))


def _shape_by_id(slide, shape_id: int):
    for index in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(index)
        if int(shape.Id) == shape_id:
            return shape
    raise ValueError(
        f"Shape ID {shape_id} was not found on slide {slide.SlideIndex}. "
        "Call ppt_list_shapes again; do not substitute a shape name or index."
    )


def _shape_text(shape) -> str:
    try:
        if shape.HasTextFrame and shape.TextFrame.HasText:
            return str(shape.TextFrame.TextRange.Text)
    except Exception:
        pass
    return ""


def _shape_state(slide_index: int, shape) -> dict:
    state = {
        "slide_index": slide_index,
        "shape_id": int(shape.Id),
        "name": str(shape.Name),
        "type": int(shape.Type),
        "left": round(float(shape.Left), 2),
        "top": round(float(shape.Top), 2),
        "width": round(float(shape.Width), 2),
        "height": round(float(shape.Height), 2),
        "rotation": round(float(shape.Rotation), 2),
        "z_order": int(shape.ZOrderPosition),
        "text": _shape_text(shape),
    }
    try:
        state["visible"] = bool(shape.Visible)
    except Exception:
        pass
    return state


def _assert_shape_precondition(shape, condition: ShapePrecondition) -> None:
    state = _shape_state(int(shape.Parent.SlideIndex), shape)
    mismatches: list[str] = []
    if condition.expected_name is not None and state["name"] != condition.expected_name:
        mismatches.append(
            f"name expected {condition.expected_name!r}, found {state['name']!r}"
        )
    for field in ("left", "top", "width", "height"):
        expected = getattr(condition, f"expected_{field}")
        if expected is not None and abs(state[field] - expected) > condition.tolerance:
            mismatches.append(f"{field} expected {expected}, found {state[field]}")
    if condition.expected_text is not None and state["text"] != condition.expected_text:
        mismatches.append(
            f"text expected {condition.expected_text!r}, found {state['text']!r}"
        )
    if mismatches:
        raise ValueError(
            f"Shape ID {state['shape_id']} precondition failed: " + "; ".join(mismatches)
        )


def _slide_indices(pres, requested: Optional[list[int]]) -> list[int]:
    if requested is None:
        return list(range(1, pres.Slides.Count + 1))
    unique = list(dict.fromkeys(requested))
    invalid = [i for i in unique if i < 1 or i > pres.Slides.Count]
    if invalid:
        raise ValueError(
            f"Slide indices out of range: {invalid}; valid range is 1-{pres.Slides.Count}"
        )
    return unique


def _capture_impl(slide_indices: Optional[list[int]]) -> dict:
    pres = ppt._get_pres_impl()
    shapes: list[dict] = []
    for slide_index in _slide_indices(pres, slide_indices):
        slide = pres.Slides(slide_index)
        for shape_index in range(1, slide.Shapes.Count + 1):
            shapes.append(_shape_state(slide_index, slide.Shapes(shape_index)))
    return {
        "presentation_full_name": str(pres.FullName),
        "slide_count": int(pres.Slides.Count),
        "shapes": shapes,
    }


def _set_work_mode_impl(params: SetWorkModeInput) -> dict:
    if params.expected_presentation_full_name:
        pres = ppt._get_pres_impl()
        actual = str(pres.FullName)
        if _normalise_windows_path(actual) != _normalise_windows_path(
            params.expected_presentation_full_name
        ):
            raise ValueError(
                f"Expected presentation '{params.expected_presentation_full_name}', "
                f"but current target is '{actual}'"
            )
        if (
            params.expected_slide_count is not None
            and int(pres.Slides.Count) != params.expected_slide_count
        ):
            raise ValueError(
                f"Expected {params.expected_slide_count} slides, "
                f"found {pres.Slides.Count}"
            )
    mode = set_work_mode(**params.model_dump())
    return {
        "success": True,
        "work_mode": {
            "allow_create": mode.allow_create,
            "require_preconditions": mode.require_preconditions,
            "expected_presentation_full_name": mode.expected_presentation_full_name,
            "expected_slide_count": mode.expected_slide_count,
        },
    }


def set_work_mode_tool(params: SetWorkModeInput) -> str:
    if params.expected_presentation_full_name:
        return json.dumps(ppt.execute(_set_work_mode_impl, params), ensure_ascii=False)
    return json.dumps(_set_work_mode_impl(params), ensure_ascii=False)


def get_work_mode_tool() -> str:
    return json.dumps({"work_mode": work_mode_dict()}, ensure_ascii=False)


def _transform_shapes_impl(params: TransformShapesInput) -> dict:
    pres = ppt._get_pres_impl()
    slide = pres.Slides(params.slide_index)
    resolved = []
    before = []
    for transform in params.transforms:
        shape = _shape_by_id(slide, transform.shape_id)
        _assert_shape_precondition(shape, transform.precondition)
        resolved.append((shape, transform))
        before.append(_shape_state(params.slide_index, shape))

    changed = []
    try:
        for shape, transform in resolved:
            for field in ("Left", "Top", "Width", "Height", "Rotation"):
                value = getattr(transform, field.lower())
                if value is not None:
                    setattr(shape, field, value)
            changed.append(shape)
    except Exception:
        for (target, _transform), original in zip(resolved, before):
            target.Left = original["left"]
            target.Top = original["top"]
            target.Width = original["width"]
            target.Height = original["height"]
            target.Rotation = original["rotation"]
        raise

    return {
        "success": True,
        "slide_index": params.slide_index,
        "before": before,
        "after": [_shape_state(params.slide_index, shape) for shape in changed],
    }


def transform_shapes(params: TransformShapesInput) -> str:
    return json.dumps(ppt.execute(_transform_shapes_impl, params), ensure_ascii=False)


def _delete_shapes_impl(params: DeleteShapesInput) -> dict:
    pres = ppt._get_pres_impl()
    slide = pres.Slides(params.slide_index)
    resolved = []
    before = []
    for shape_id, precondition in params.targets.items():
        shape = _shape_by_id(slide, shape_id)
        _assert_shape_precondition(shape, precondition)
        resolved.append(shape)
        before.append(_shape_state(params.slide_index, shape))

    # Resolve and validate every target before the first destructive call.
    for shape in sorted(resolved, key=lambda item: item.ZOrderPosition, reverse=True):
        shape.Delete()

    return {
        "success": True,
        "slide_index": params.slide_index,
        "deleted": before,
        "deleted_shape_ids": [item["shape_id"] for item in before],
        "created_shape_ids": [],
    }


def delete_shapes(params: DeleteShapesInput) -> str:
    return json.dumps(ppt.execute(_delete_shapes_impl, params), ensure_ascii=False)


def _replacement_bbox(shapes) -> tuple[float, float, float, float]:
    left = min(float(shape.Left) for shape in shapes)
    top = min(float(shape.Top) for shape in shapes)
    right = max(float(shape.Left + shape.Width) for shape in shapes)
    bottom = max(float(shape.Top + shape.Height) for shape in shapes)
    return left, top, right - left, bottom - top


def _replace_visual_impl(params: ReplaceVisualInput) -> dict:
    pres = ppt._get_pres_impl()
    slide = pres.Slides(params.slide_index)
    targets = [_shape_by_id(slide, shape_id) for shape_id in params.target_shape_ids]
    replacement = _shape_by_id(slide, params.replacement_shape_id)
    for shape in targets:
        _assert_shape_precondition(shape, params.target_preconditions[int(shape.Id)])
    _assert_shape_precondition(replacement, params.replacement_precondition)

    target_before = [_shape_state(params.slide_index, shape) for shape in targets]
    replacement_before = _shape_state(params.slide_index, replacement)
    left, top, width, height = _replacement_bbox(targets)

    replacement.Left = left
    replacement.Top = top
    if params.fit_mode == "fill":
        replacement.Width = width
        replacement.Height = height
    elif params.fit_mode == "contain":
        source_ratio = replacement_before["width"] / replacement_before["height"]
        target_ratio = width / height
        if source_ratio >= target_ratio:
            replacement.Width = width
            replacement.Height = width / source_ratio
            replacement.Top = top + (height - replacement.Height) / 2
        else:
            replacement.Height = height
            replacement.Width = height * source_ratio
            replacement.Left = left + (width - replacement.Width) / 2

    deleted = []
    if params.delete_targets:
        deleted = [int(shape.Id) for shape in targets]
        for shape in reversed(targets):
            shape.Delete()

    return {
        "success": True,
        "slide_index": params.slide_index,
        "replacement_before": replacement_before,
        "replacement_after": _shape_state(params.slide_index, replacement),
        "replaced_targets": target_before,
        "deleted_shape_ids": deleted,
        "created_shape_ids": [],
    }


def replace_visual(params: ReplaceVisualInput) -> str:
    return json.dumps(ppt.execute(_replace_visual_impl, params), ensure_ascii=False)


def capture_shape_snapshot(params: SnapshotInput) -> str:
    snapshot = ppt.execute(_capture_impl, params.slide_indices)
    snapshot_id = uuid.uuid4().hex
    with _snapshot_lock:
        _snapshots[snapshot_id] = snapshot
        _snapshots.move_to_end(snapshot_id)
        while len(_snapshots) > _MAX_SNAPSHOTS:
            _snapshots.popitem(last=False)
    return json.dumps(
        {
            "success": True,
            "snapshot_id": snapshot_id,
            "presentation_full_name": snapshot["presentation_full_name"],
            "slide_count": snapshot["slide_count"],
            "shape_count": len(snapshot["shapes"]),
        },
        ensure_ascii=False,
    )


def _state_map(snapshot: dict) -> dict[tuple[int, int], dict]:
    return {
        (shape["slide_index"], shape["shape_id"]): shape
        for shape in snapshot["shapes"]
    }


def diff_shape_snapshot(params: DiffSnapshotInput) -> str:
    with _snapshot_lock:
        before = _snapshots.get(params.snapshot_id)
    if before is None:
        raise ValueError(
            f"Snapshot '{params.snapshot_id}' was not found or expired. "
            "Capture a new snapshot before editing."
        )
    current = ppt.execute(_capture_impl, None)
    if _normalise_windows_path(current["presentation_full_name"]) != _normalise_windows_path(
        before["presentation_full_name"]
    ):
        raise ValueError("The active presentation does not match the snapshot")

    old_map = _state_map(before)
    new_map = _state_map(current)
    added = [new_map[key] for key in sorted(new_map.keys() - old_map.keys())]
    deleted = [old_map[key] for key in sorted(old_map.keys() - new_map.keys())]
    modified = []
    ignored = {"z_order"}
    for key in sorted(old_map.keys() & new_map.keys()):
        old = old_map[key]
        new = new_map[key]
        fields = [
            field
            for field in old.keys() | new.keys()
            if field not in ignored and old.get(field) != new.get(field)
        ]
        if fields:
            modified.append(
                {
                    "slide_index": key[0],
                    "shape_id": key[1],
                    "changed_fields": sorted(fields),
                    "before": old,
                    "after": new,
                }
            )
    return json.dumps(
        {
            "success": True,
            "snapshot_id": params.snapshot_id,
            "added": added,
            "deleted": deleted,
            "modified": modified,
            "summary": {
                "added": len(added),
                "deleted": len(deleted),
                "modified": len(modified),
            },
        },
        ensure_ascii=False,
    )


def _minimum_font_size(shape) -> Optional[float]:
    try:
        if not shape.HasTextFrame or not shape.TextFrame.HasText:
            return None
        text_range = shape.TextFrame.TextRange
        size = float(text_range.Font.Size)
        if size > 0:
            return size
        sizes = []
        for index in range(1, int(text_range.Length) + 1):
            char_size = float(text_range.Characters(index, 1).Font.Size)
            if char_size > 0:
                sizes.append(char_size)
        return min(sizes) if sizes else None
    except Exception:
        return None


def _validate_impl(params: ValidatePresentationInput) -> dict:
    pres = ppt._get_pres_impl()
    findings = []
    checked_shapes = 0
    for slide_index in _slide_indices(pres, params.slide_indices):
        slide = pres.Slides(slide_index)
        for shape_index in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(shape_index)
            checked_shapes += 1
            state = _shape_state(slide_index, shape)
            font_size = _minimum_font_size(shape)
            if font_size is not None and font_size < params.min_font_size:
                findings.append(
                    {
                        "code": "FONT_TOO_SMALL",
                        "slide_index": slide_index,
                        "shape_id": state["shape_id"],
                        "shape_name": state["name"],
                        "font_size": round(font_size, 2),
                        "minimum": params.min_font_size,
                    }
                )
            if (
                state["width"] < params.minimum_frame_size
                or state["height"] < params.minimum_frame_size
            ):
                findings.append(
                    {
                        "code": "TINY_FRAME",
                        "slide_index": slide_index,
                        "shape_id": state["shape_id"],
                        "shape_name": state["name"],
                        "width": state["width"],
                        "height": state["height"],
                    }
                )
            if state["type"] in (14, 17) and not state["text"].strip():
                findings.append(
                    {
                        "code": "EMPTY_TEXT_FRAME",
                        "slide_index": slide_index,
                        "shape_id": state["shape_id"],
                        "shape_name": state["name"],
                    }
                )
    return {
        "valid": not findings,
        "presentation_full_name": str(pres.FullName),
        "checked_shapes": checked_shapes,
        "findings": findings,
        "summary": {
            "finding_count": len(findings),
            "by_code": {
                code: sum(1 for finding in findings if finding["code"] == code)
                for code in sorted({finding["code"] for finding in findings})
            },
        },
    }


def validate_presentation(params: ValidatePresentationInput) -> str:
    return json.dumps(ppt.execute(_validate_impl, params), ensure_ascii=False)


def register_tools(mcp):
    @mcp.tool(
        name="ppt_set_work_mode",
        annotations={
            "title": "Set Safe Work Mode",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_set_work_mode(params: SetWorkModeInput) -> str:
        """Set creation policy and presentation-level edit preconditions."""
        return set_work_mode_tool(params)

    @mcp.tool(
        name="ppt_get_work_mode",
        annotations={
            "title": "Get Safe Work Mode",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_get_work_mode() -> str:
        """Return the current creation and precondition policy."""
        return get_work_mode_tool()

    @mcp.tool(
        name="ppt_transform_shapes",
        annotations={
            "title": "Transform Existing Shapes by ID",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_transform_shapes(params: TransformShapesInput) -> str:
        """Atomically move, resize, or rotate existing shapes after checking observed state."""
        return transform_shapes(params)

    @mcp.tool(
        name="ppt_replace_visual",
        annotations={
            "title": "Replace Visual with Existing Shape",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_replace_visual(params: ReplaceVisualInput) -> str:
        """Replace target shapes with an authentic existing shape; never generates a lookalike."""
        return replace_visual(params)

    @mcp.tool(
        name="ppt_delete_shapes",
        annotations={
            "title": "Delete Existing Shapes by ID",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_delete_shapes(params: DeleteShapesInput) -> str:
        """Delete existing shapes only after validating their observed state."""
        return delete_shapes(params)

    @mcp.tool(
        name="ppt_capture_shape_snapshot",
        annotations={
            "title": "Capture Shape Snapshot",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    async def tool_capture_shape_snapshot(params: SnapshotInput) -> str:
        """Capture stable shape IDs and properties before editing."""
        return capture_shape_snapshot(params)

    @mcp.tool(
        name="ppt_diff_shape_snapshot",
        annotations={
            "title": "Diff Shape Snapshot",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_diff_shape_snapshot(params: DiffSnapshotInput) -> str:
        """Report added, deleted, and modified shapes relative to a snapshot."""
        return diff_shape_snapshot(params)

    @mcp.tool(
        name="ppt_validate_presentation",
        annotations={
            "title": "Validate Presentation Quality",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    async def tool_validate_presentation(params: ValidatePresentationInput) -> str:
        """Find undersized text, tiny frames, and empty text placeholders."""
        return validate_presentation(params)

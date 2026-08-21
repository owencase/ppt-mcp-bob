"""Tests for IBM Bob safety policy and stable shape IDs."""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from pydantic import BaseModel, ValidationError

from ppt_com.safe_ops import (
    DeleteShapesInput,
    SetWorkModeInput,
    ShapePrecondition,
    ShapeTransform,
    TransformShapesInput,
)
from ppt_com.safety import set_work_mode
from ppt_com.shapes import ShapeIdentifierInput, _get_shape
from utils.bob_safety import (
    BobToolError,
    make_strict_input_models,
    promote_error_result,
    validate_session_preconditions,
)


class StrictTestInput(BaseModel):
    value: int


async def _strict_test_tool(params: StrictTestInput):
    return params.value


def test_creation_is_blocked_by_default():
    set_work_mode(allow_create=False, require_preconditions=False)
    with pytest.raises(BobToolError) as exc:
        validate_session_preconditions("ppt_add_shape", mutating=True)
    assert exc.value.payload["retryable"] is False
    assert "ppt_add" in exc.value.payload["hint"]


def test_direct_shape_mutation_is_blocked_in_precondition_mode(monkeypatch):
    set_work_mode(
        allow_create=False,
        require_preconditions=True,
        expected_presentation_full_name=r"C:\\Decks\\target.pptx",
        expected_slide_count=10,
    )
    monkeypatch.setattr(
        "utils.bob_safety._target_state_impl",
        lambda: {
            "full_name": r"C:\\Decks\\target.pptx",
            "name": "target.pptx",
            "presentation_index": 2,
            "slide_count": 10,
        },
    )
    monkeypatch.setattr("utils.com_wrapper.ppt.execute", lambda func: func())

    with pytest.raises(BobToolError) as exc:
        validate_session_preconditions("ppt_update_shape", mutating=True)

    assert exc.value.payload["retryable"] is False
    assert "ppt_transform_shapes" in exc.value.payload["hint"]


def test_explicit_selector_cannot_bypass_locked_target(monkeypatch):
    set_work_mode(
        allow_create=False,
        require_preconditions=True,
        expected_presentation_full_name=r"C:\\Decks\\target.pptx",
        expected_slide_count=10,
    )
    monkeypatch.setattr(
        "utils.bob_safety._target_state_impl",
        lambda: {
            "full_name": r"C:\\Decks\\target.pptx",
            "name": "target.pptx",
            "presentation_index": 2,
            "slide_count": 10,
        },
    )
    monkeypatch.setattr("utils.com_wrapper.ppt.execute", lambda func: func())

    with pytest.raises(BobToolError) as exc:
        validate_session_preconditions(
            "ppt_delete_slide",
            mutating=True,
            params=SimpleNamespace(presentation_index=1),
        )

    assert "locked target" in exc.value.payload["error"]


def test_legacy_error_json_is_promoted_to_retry_aware_failure():
    with pytest.raises(BobToolError) as exc:
        promote_error_result("ppt_update_shape", json.dumps({"error": "missing"}))
    assert exc.value.payload == {
        "error": "ppt_update_shape: missing",
        "retryable": False,
        "hint": "Correct the input or inspect the target state before retrying.",
    }


def test_work_mode_requires_expected_file_in_safe_mode():
    with pytest.raises(ValidationError):
        SetWorkModeInput(require_preconditions=True)


def test_transform_requires_observed_shape_state():
    with pytest.raises(ValidationError):
        ShapePrecondition()
    transform = ShapeTransform(
        shape_id=101,
        precondition=ShapePrecondition(expected_name="ObservedShape"),
        left=100,
    )
    params = TransformShapesInput(slide_index=1, transforms=[transform])
    assert params.transforms[0].shape_id == 101


def test_safe_delete_requires_shape_id_and_observed_state():
    params = DeleteShapesInput(
        slide_index=1,
        targets={101: ShapePrecondition(expected_name="ObservedShape")},
    )
    assert list(params.targets) == [101]
    with pytest.raises(ValidationError):
        DeleteShapesInput(
            slide_index=1,
            targets={0: {"expected_name": "ObservedShape"}},
        )


def test_shape_identifier_accepts_stable_id_only():
    assert ShapeIdentifierInput(slide_index=1, shape_id=101).shape_id == 101
    with pytest.raises(ValidationError):
        ShapeIdentifierInput(
            slide_index=1,
            shape_id=101,
            shape_name="ObservedShape",
        )


def test_get_shape_resolves_shape_id_not_collection_index():
    class Shape:
        def __init__(self, shape_id, name):
            self.Id = shape_id
            self.Name = name

    class Shapes:
        def __init__(self):
            self.items = [Shape(42, "first"), Shape(7, "second")]
            self.Count = len(self.items)

        def __call__(self, index):
            return self.items[index - 1]

    slide = type("Slide", (), {})()
    slide.Shapes = Shapes()
    assert _get_shape(slide, None, shape_id=7).Name == "second"


def test_registered_input_models_forbid_unknown_fields():
    make_strict_input_models(_strict_test_tool)
    with pytest.raises(ValidationError):
        StrictTestInput(value=1, ignored="must fail")

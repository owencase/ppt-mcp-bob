"""Tests for React Bits-inspired PowerPoint-native design references."""

from __future__ import annotations

import json

import pytest
from pydantic import ValidationError

from tools.design_reference import GetDesignReferenceInput, get_design_reference


def _payload(**kwargs):
    return json.loads(get_design_reference(GetDesignReferenceInput(**kwargs)))


def test_lists_curated_reference_styles_without_fetching_code():
    payload = _payload()

    assert payload["success"] is True
    assert payload["reference_only"] is True
    assert len(payload["styles"]) == 7
    assert all(item["reference_url"].startswith("https://") for item in payload["styles"])


def test_agency_cover_returns_native_design_contract():
    payload = _payload(
        style="agency_dark",
        slide_purpose="cover",
        motion="subtle",
    )

    assert payload["mode"] == "powerpoint_native_redesign"
    assert payload["reference"]["name"] == "Agency Dark"
    assert payload["design_contract"]["canvas"] == {
        "ratio": "16:9",
        "width_pt": 960,
        "height_pt": 540,
    }
    assert payload["design_contract"]["layout"]["safe_margin_pt"] == 54
    assert "ppt_add_animation" in payload["recommended_tools"]


def test_no_motion_returns_no_object_effects():
    payload = _payload(style="glass_surface", motion="none")

    assert payload["design_contract"]["motion"]["effects"] == []
    assert payload["design_contract"]["motion"]["transition"] == {"effect": "none"}


def test_recipe_explicitly_rejects_react_runtime_translation():
    payload = _payload(style="aurora_gradient")
    guardrails = " ".join(payload["guardrails"])

    assert "Do not import or execute React Bits code" in guardrails
    assert "WebGL" in guardrails


def test_unknown_style_is_rejected_by_schema():
    with pytest.raises(ValidationError):
        GetDesignReferenceInput(style="unknown")


def test_unknown_field_is_rejected_by_schema():
    with pytest.raises(ValidationError):
        GetDesignReferenceInput(style="agency_dark", url="https://example.com")

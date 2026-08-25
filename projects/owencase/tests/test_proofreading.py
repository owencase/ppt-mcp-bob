"""Tests for location-aware presentation text proofreading."""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ppt_com import proofreading
from ppt_com.constants import msoChart, msoGroup, msoSmartArt, msoTable, msoTextBox
from ppt_com.proofreading import (
    ProofreadTextInput,
    _analyse_text_units,
    _collect_text_units_impl,
    proofread_text,
)


def _unit(text: str, **location):
    return {
        "unit_id": "slide-1:unit-1",
        "slide_index": 1,
        "container_type": "shape_text",
        "shape_id": 7,
        "shape_name": "Title 1",
        "text": text,
        **location,
    }


def _only_rules(**overrides):
    values = {
        "check_common_typos": False,
        "check_repeated_words": False,
        "check_punctuation": False,
        "check_brackets": False,
        "check_suspicious_characters": False,
    }
    values.update(overrides)
    return ProofreadTextInput(**values)


class TestProofreadTextInput:
    def test_defaults_return_text_for_contextual_review(self):
        params = ProofreadTextInput()
        assert params.include_text_units is True
        assert params.include_notes is False
        assert params.max_findings == 200

    @pytest.mark.parametrize("indices", [[], [0], [1, 1]])
    def test_invalid_slide_indices_are_rejected(self, indices):
        with pytest.raises(ValidationError):
            ProofreadTextInput(slide_indices=indices)

    def test_empty_or_noop_custom_replacements_are_rejected(self):
        with pytest.raises(ValidationError):
            ProofreadTextInput(custom_replacements={"": "고객"})
        with pytest.raises(ValidationError):
            ProofreadTextInput(custom_replacements={"고객": "고객"})

    def test_duplicate_allowed_terms_are_rejected_case_insensitively(self):
        with pytest.raises(ValidationError):
            ProofreadTextInput(allowed_terms=["OpenAI", "openai"])


class TestDeterministicProofreading:
    def test_common_korean_typo_has_location_and_suggestion(self):
        findings, total = _analyse_text_units(
            [_unit("신규 프리젠테이션 전략")], ProofreadTextInput()
        )

        typo = next(item for item in findings if item["code"] == "SUSPECTED_TYPO")
        assert total == 1
        assert typo["original"] == "프리젠테이션"
        assert typo["suggestion"] == "프레젠테이션"
        assert typo["slide_index"] == 1
        assert typo["shape_id"] == 7
        assert typo["line"] == 1
        assert typo["column"] == 4

    def test_custom_replacement_finds_domain_typo(self):
        params = _only_rules(custom_replacements={"거겍": "고객"})
        findings, total = _analyse_text_units([_unit("핵심 거겍 분석")], params)

        assert total == 1
        assert findings[0]["source"] == "custom"
        assert findings[0]["suggestion"] == "고객"

    def test_allowed_term_suppresses_built_in_replacement(self):
        params = ProofreadTextInput(allowed_terms=["컨텐츠"])
        findings, total = _analyse_text_units([_unit("컨텐츠 전략")], params)

        assert findings == []
        assert total == 0

    def test_repeated_words_brackets_punctuation_and_bad_character(self):
        text = "전략 전략 추진  (2026년!!\uFFFD"
        findings, total = _analyse_text_units([_unit(text)], ProofreadTextInput())
        codes = {item["code"] for item in findings}

        assert total >= 5
        assert {
            "REPEATED_WORD",
            "UNMATCHED_BRACKET",
            "SUSPICIOUS_PUNCTUATION",
            "SUSPICIOUS_CHARACTER",
        } <= codes

    def test_clean_text_passes_deterministic_checks(self):
        findings, total = _analyse_text_units(
            [_unit("고객 성장을 위한 데이터 전략입니다.")], ProofreadTextInput()
        )

        assert findings == []
        assert total == 0

    def test_findings_are_capped_but_total_is_preserved(self):
        params = _only_rules(
            check_repeated_words=True,
            max_findings=1,
        )
        findings, total = _analyse_text_units(
            [_unit("고객 고객 전략 전략 성과 성과")], params
        )

        assert len(findings) == 1
        assert total == 3

    def test_table_cell_keeps_cell_and_text_columns(self):
        unit = _unit(
            "거겍 분석",
            container_type="table_cell",
            row=2,
            column=3,
        )
        params = _only_rules(custom_replacements={"거겍": "고객"})
        findings, _ = _analyse_text_units([unit], params)

        assert findings[0]["row"] == 2
        assert findings[0]["column"] == 3
        assert findings[0]["text_column"] == 1


class _Collection:
    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


def _text_shape(name, shape_id, text):
    return SimpleNamespace(
        Type=msoTextBox,
        Name=name,
        Id=shape_id,
        HasTextFrame=True,
        TextFrame=SimpleNamespace(
            HasText=bool(text),
            TextRange=SimpleNamespace(Text=text),
        ),
    )


class _FakeTable:
    def __init__(self, values):
        self._values = values
        self.Rows = SimpleNamespace(Count=len(values))
        self.Columns = SimpleNamespace(Count=len(values[0]))

    def Cell(self, row, column):
        return SimpleNamespace(
            Shape=SimpleNamespace(
                TextFrame=SimpleNamespace(
                    TextRange=SimpleNamespace(Text=self._values[row - 1][column - 1])
                )
            )
        )


class _FakeSeriesCollection:
    def __init__(self, series):
        self._series = series
        self.Count = len(series)

    def __call__(self, index):
        return self._series[index - 1]


class _FakeChart:
    HasTitle = True
    ChartTitle = SimpleNamespace(Text="차트 메세지")

    def __init__(self):
        self._series = _FakeSeriesCollection(
            [SimpleNamespace(Name="매출 시리즈", XValues=("상반기", "하반기"))]
        )

    def SeriesCollection(self, index=None):
        if index is None:
            return self._series
        return self._series(index)

    def Axes(self, axis_type):
        title = "분기" if axis_type == 1 else "매출"
        return SimpleNamespace(HasTitle=True, AxisTitle=SimpleNamespace(Text=title))


def _fake_presentation():
    regular = _text_shape("Body", 1, "일반 텍스트")
    group_child = _text_shape("Grouped Text", 3, "그룹 텍스트")
    group = SimpleNamespace(
        Type=msoGroup,
        Name="Group 1",
        Id=2,
        GroupItems=_Collection([group_child]),
    )
    table = SimpleNamespace(
        Type=msoTable,
        Name="Table 1",
        Id=4,
        Table=_FakeTable([["표 제목", "표 데이터"]]),
    )
    nodes = _Collection(
        [SimpleNamespace(TextFrame2=SimpleNamespace(TextRange=SimpleNamespace(Text="스마트아트")))]
    )
    smartart = SimpleNamespace(
        Type=msoSmartArt,
        Name="SmartArt 1",
        Id=5,
        SmartArt=SimpleNamespace(AllNodes=nodes),
    )
    chart = SimpleNamespace(
        Type=msoChart,
        Name="Chart 1",
        Id=6,
        Chart=_FakeChart(),
    )
    notes = SimpleNamespace(
        Shapes=SimpleNamespace(
            Placeholders=_Collection(
                [
                    SimpleNamespace(),
                    SimpleNamespace(
                        TextFrame=SimpleNamespace(
                            TextRange=SimpleNamespace(Text="발표자 노트")
                        )
                    ),
                ]
            )
        )
    )
    slide = SimpleNamespace(
        Shapes=_Collection([regular, group, table, smartart, chart]),
        NotesPage=notes,
    )
    return SimpleNamespace(
        FullName=r"C:\demo\proofread.pptx",
        Slides=_Collection([slide]),
    )


class TestTextCollection:
    def test_collects_all_supported_visible_container_types(self, monkeypatch):
        presentation = _fake_presentation()
        monkeypatch.setattr(proofreading.ppt, "_get_pres_impl", lambda: presentation)

        units = _collect_text_units_impl([1], include_notes=True)
        container_types = {unit["container_type"] for unit in units}

        assert {
            "shape_text",
            "table_cell",
            "smartart_node",
            "chart_title",
            "chart_axis_title",
            "chart_series_name",
            "chart_category",
            "speaker_notes",
        } <= container_types
        grouped = next(unit for unit in units if unit["text"] == "그룹 텍스트")
        assert grouped["group_path"] == ["Group 1"]

    def test_full_result_requires_context_review_and_detects_chart_typo(self, monkeypatch):
        presentation = _fake_presentation()
        monkeypatch.setattr(proofreading.ppt, "_get_pres_impl", lambda: presentation)
        monkeypatch.setattr(
            proofreading.ppt,
            "execute",
            lambda function, *args: function(*args),
        )

        result = json.loads(proofread_text(ProofreadTextInput(include_notes=True)))

        assert result["valid"] is False
        assert result["semantic_review_required"] is True
        assert result["checked_text_units"] == len(result["text_units"])
        assert any(
            item["container_type"] == "chart_title"
            and item.get("suggestion") == "메시지"
            for item in result["findings"]
        )

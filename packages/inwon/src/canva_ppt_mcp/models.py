from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field, field_validator, model_validator


LayoutName = Literal[
    "title", "two_column", "icon_rows", "big_stat", "grid_2x2",
    "timeline", "comparison", "image_focus", "chart", "closing",
]


def _validate_hex(value: str) -> str:
    value = value.upper()
    if len(value) != 7 or not value.startswith("#"):
        raise ValueError("Color must be #RRGGBB")
    int(value[1:], 16)
    return value


class Palette(BaseModel):
    primary: str
    secondary: list[str] = Field(min_length=1, max_length=2)
    accent: str
    background_light: str
    background_dark: str | None = None

    @field_validator("primary", "accent", "background_light", "background_dark")
    @classmethod
    def valid_hex(cls, value: str | None) -> str | None:
        if value is None:
            return value
        return _validate_hex(value)

    @field_validator("secondary")
    @classmethod
    def valid_secondary_hex(cls, values: list[str]) -> list[str]:
        return [_validate_hex(value) for value in values]


class Typography(BaseModel):
    header_font: Literal["Cambria", "Bookman Old Style", "Century Schoolbook"] = "Cambria"
    body_font: Literal["Calibri", "Arial"] = "Arial"
    title_size: int = Field(default=52, ge=44, le=60)
    section_header_size: int = Field(default=22, ge=20, le=28)
    body_size: int = Field(default=16, ge=14, le=20)
    caption_size: int = Field(default=11, ge=10, le=12)


class DesignSystem(BaseModel):
    palette: Palette
    typography: Typography
    visual_motif: str
    style_preset: Literal["orbital", "editorial", "neon", "organic", "luxury", "geometric", "swiss"] = "orbital"
    layout_rotation: list[LayoutName] = Field(min_length=4)
    visual_intensity: Literal["balanced", "bold", "maximal"] = "bold"
    dark_slide_ratio: float = Field(default=.45, ge=.40, le=.50)
    gradient_backgrounds: bool = True
    background_texture: bool = True
    dynamic_composition: bool = True

    @field_validator("layout_rotation")
    @classmethod
    def valid_rotation(cls, values: list[LayoutName]) -> list[LayoutName]:
        if len(set(values)) != len(values):
            raise ValueError("layout_rotation must not contain duplicates")
        if len([x for x in values if x not in {"title", "closing"}]) < 2:
            raise ValueError("layout_rotation needs at least two body layouts")
        return values


class ContentItem(BaseModel):
    heading: str = ""
    body: str = ""
    value: str | None = None
    image_path: str | None = None
    image_url: str | None = None
    claim_ids: list[str] = Field(default_factory=list)


class ChartSeries(BaseModel):
    name: str
    values: list[float] = Field(min_length=1, max_length=12)


class ChartSpec(BaseModel):
    chart_type: Literal["bar", "column", "line", "pie"] = "column"
    categories: list[str] = Field(min_length=1, max_length=12)
    series: list[ChartSeries] = Field(min_length=1, max_length=4)
    value_suffix: str = ""

    @model_validator(mode="after")
    def series_lengths_match(self):
        if any(len(series.values) != len(self.categories) for series in self.series):
            raise ValueError("every chart series must match categories length")
        return self


class SlideSpec(BaseModel):
    title: str
    subtitle: str = ""
    layout: LayoutName
    items: list[ContentItem] = Field(default_factory=list, max_length=6)
    chart: ChartSpec | None = None
    speaker_notes: str = ""


class ResearchSource(BaseModel):
    title: str
    url: str
    source_type: Literal["wikipedia", "web", "user", "asset"] = "web"
    retrieved_at: str | None = None


class EvidenceClaim(BaseModel):
    claim_id: str
    text: str
    source_url: str
    section: str = ""
    numeric_values: list[str] = Field(default_factory=list)


class DeckPlan(BaseModel):
    communication_job: str
    design_system: DesignSystem
    slides: list[SlideSpec] = Field(min_length=2, max_length=30)
    research_sources: list[ResearchSource] = Field(default_factory=list)
    evidence_claims: list[EvidenceClaim] = Field(default_factory=list)
    grounded: bool = False
    language: Literal["ko", "en"] = "ko"


class QAIssue(BaseModel):
    slide: int
    code: str
    severity: Literal["warning", "error"] = "error"
    message: str
    shape_name: str | None = None


class QAReport(BaseModel):
    passed: bool
    rounds: int
    issues: list[QAIssue] = Field(default_factory=list)
    rendered_slides: list[str] = Field(default_factory=list)

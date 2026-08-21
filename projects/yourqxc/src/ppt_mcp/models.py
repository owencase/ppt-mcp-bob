"""LLM이 채우는 덱/슬라이드 스펙.

이 모듈의 필드 설명(description)이 그대로 MCP 도구의 JSON 스키마가 되므로,
에이전트가 읽고 바로 채울 수 있게 구체적으로 적는다.
"""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, model_validator

ColorToken = Literal["default", "muted", "accent", "accent_alt", "positive", "negative"]
Align = Literal["left", "center", "right"]


class _Base(BaseModel):
    model_config = ConfigDict(extra="forbid")


# --- 조각 -----------------------------------------------------------------

class Bullet(_Base):
    """불릿 한 줄. 문자열로도 쓸 수 있다(앞의 '-'/'*' 개수가 들여쓰기 단계)."""

    text: str = Field(description="불릿 본문. 한 줄에 한 가지 메시지만 담는 편이 읽기 좋다.")
    level: int = Field(0, ge=0, le=3, description="들여쓰기 단계. 0이 최상위.")
    bold: bool = Field(False, description="굵게 표시할지 여부.")
    color: ColorToken = Field("default", description="테마 색 토큰. 강조는 accent, 증감은 positive/negative.")


BulletLike = Annotated[
    Union[str, Bullet],
    Field(union_mode="left_to_right",
          description="문자열(예: '핵심 지표', '- 하위 항목') 또는 Bullet 객체."),
]


class Column(_Base):
    """2단 레이아웃의 한쪽 열."""

    heading: str | None = Field(None, description="열 제목. 없으면 본문만 표시.")
    bullets: list[BulletLike] = Field(default_factory=list, description="열 안의 불릿 목록.")
    text: str | None = Field(None, description="불릿 대신 문단으로 넣고 싶을 때 사용.")


class ImageSource(_Base):
    """이미지 입력. 세 가지 중 정확히 하나만 채운다."""

    path: str | None = Field(None, description="로컬 이미지 파일 경로.")
    base64: str | None = Field(None, description="base64로 인코딩한 이미지 바이트(data URI 접두사 허용).")
    url: str | None = Field(None, description="이미지 URL. PPT_MCP_ALLOW_REMOTE_IMAGES=1일 때만 허용.")

    @model_validator(mode="after")
    def _exactly_one(self) -> "ImageSource":
        filled = [n for n in ("path", "base64", "url") if getattr(self, n)]
        if len(filled) != 1:
            raise ValueError(f"path/base64/url 중 정확히 하나만 채워야 합니다(현재: {filled or '없음'}).")
        return self


class Series(_Base):
    """차트 계열 하나."""

    name: str = Field(description="범례에 표시될 계열 이름.")
    values: list[float] = Field(description="categories와 길이가 같아야 하는 값 목록.")


class KpiItem(_Base):
    """숫자 강조 카드 하나."""

    value: str = Field(description="크게 보여줄 값. 예: '42%', '1,204건', '3.2배'.")
    label: str = Field(description="값 아래 설명 문구.")
    delta: str | None = Field(None, description="증감 표기. 예: '+12%p', '-3건'.")
    tone: Literal["neutral", "positive", "negative"] = Field(
        "neutral", description="delta 색상. 개선이면 positive, 악화면 negative.")


class TimelineStep(_Base):
    """타임라인 한 단계."""

    label: str = Field(description="단계 표시. 예: '1분기', 'Phase 1', '8/20'.")
    title: str = Field(description="단계 제목.")
    description: str | None = Field(None, description="한 줄 보충 설명.")


class TextBlock(_Base):
    """blank 슬라이드에서 위치를 직접 지정하는 텍스트 상자. 단위는 인치."""

    text: str = Field(description="표시할 텍스트. 줄바꿈은 \\n.")
    x: float = Field(description="왼쪽 위치(인치). 슬라이드 폭은 13.333.")
    y: float = Field(description="위쪽 위치(인치). 슬라이드 높이는 7.5.")
    w: float = Field(description="너비(인치).")
    h: float = Field(description="높이(인치).")
    size: int | None = Field(None, description="글자 크기(pt). 생략하면 본문 크기.")
    bold: bool = Field(False, description="굵게 표시할지 여부.")
    align: Align = Field("left", description="가로 정렬.")
    color: ColorToken = Field("default", description="테마 색 토큰.")


# --- 슬라이드 --------------------------------------------------------------

class _Slide(_Base):
    notes: str | None = Field(None, description="발표자 노트. 슬라이드에는 보이지 않는다.")
    layout: str | None = Field(
        None, description="템플릿 사용 시 강제로 지정할 레이아웃 이름. 보통 비워 둔다.")


class TitleSlide(_Slide):
    """표지."""

    type: Literal["title"] = "title"
    title: str = Field(description="덱 제목.")
    subtitle: str | None = Field(None, description="부제 한 줄.")
    eyebrow: str | None = Field(None, description="제목 위 작은 라벨. 예: '주간 보고', 'IBM Korea'.")
    presenter: str | None = Field(None, description="발표자/작성자.")
    date: str | None = Field(None, description="날짜 문자열. 예: '2026-08-20'.")


class AgendaSlide(_Slide):
    """번호가 매겨진 목차."""

    type: Literal["agenda"] = "agenda"
    title: str = Field("목차", description="슬라이드 제목.")
    items: list[str] = Field(description="목차 항목. 4~7개가 적당하다.")


class SectionSlide(_Slide):
    """장 구분용 간지."""

    type: Literal["section"] = "section"
    title: str = Field(description="섹션 제목.")
    subtitle: str | None = Field(None, description="섹션 한 줄 요약.")
    number: str | None = Field(None, description="섹션 번호. 예: '01'.")


class BulletsSlide(_Slide):
    """제목 + 불릿. 가장 많이 쓰는 형식."""

    type: Literal["bullets"] = "bullets"
    title: str = Field(description="슬라이드 제목.")
    subtitle: str | None = Field(None, description="제목 아래 한 줄 설명.")
    bullets: list[BulletLike] = Field(description="불릿 목록. 한 장에 6개 이하를 권장.")


class TwoColumnSlide(_Slide):
    """좌우 2단 구성."""

    type: Literal["two_column"] = "two_column"
    title: str = Field(description="슬라이드 제목.")
    subtitle: str | None = Field(None, description="제목 아래 한 줄 설명.")
    left: Column = Field(description="왼쪽 열.")
    right: Column = Field(description="오른쪽 열.")


class ComparisonSlide(_Slide):
    """색이 구분된 카드 두 장으로 비교(장단점, As-Is/To-Be 등)."""

    type: Literal["comparison"] = "comparison"
    title: str = Field(description="슬라이드 제목.")
    left: Column = Field(description="왼쪽 카드. heading에 'As-Is' 같은 라벨.")
    right: Column = Field(description="오른쪽 카드. heading에 'To-Be' 같은 라벨.")
    left_tone: Literal["neutral", "positive", "negative"] = Field(
        "negative", description="왼쪽 카드 강조색.")
    right_tone: Literal["neutral", "positive", "negative"] = Field(
        "positive", description="오른쪽 카드 강조색.")


class TableSlide(_Slide):
    """표."""

    type: Literal["table"] = "table"
    title: str = Field(description="슬라이드 제목.")
    headers: list[str] = Field(description="머리글 행.")
    rows: list[list[str]] = Field(description="데이터 행. 각 행의 길이는 headers와 같아야 한다.")
    column_widths: list[float] | None = Field(
        None, description="열 너비 비율. 예: [2,1,1]. 생략하면 균등 분할.")
    caption: str | None = Field(None, description="표 아래 출처/각주.")

    @model_validator(mode="after")
    def _check_shape(self) -> "TableSlide":
        width = len(self.headers)
        for i, row in enumerate(self.rows):
            if len(row) != width:
                raise ValueError(f"{i}번째 행의 칸 수({len(row)})가 headers({width})와 다릅니다.")
        if self.column_widths and len(self.column_widths) != width:
            raise ValueError(f"column_widths 길이({len(self.column_widths)})가 headers({width})와 다릅니다.")
        return self


class ChartSlide(_Slide):
    """차트."""

    type: Literal["chart"] = "chart"
    title: str = Field(description="슬라이드 제목.")
    chart_type: Literal["column", "bar", "line", "pie", "doughnut", "area", "scatter"] = Field(
        "column", description="차트 종류. 시간 흐름은 line, 구성비는 pie/doughnut, 항목 비교는 column.")
    categories: list[str] = Field(description="x축 항목(또는 파이 조각 이름).")
    series: list[Series] = Field(description="계열 목록. pie/doughnut은 계열 1개만 사용한다.")
    caption: str | None = Field(None, description="차트 아래 출처/각주.")
    show_legend: bool = Field(True, description="범례 표시 여부. 계열이 1개면 끄는 편이 깔끔하다.")
    show_data_labels: bool = Field(False, description="데이터 레이블 표시 여부.")

    @model_validator(mode="after")
    def _check_lengths(self) -> "ChartSlide":
        if not self.series:
            raise ValueError("series가 비어 있습니다.")
        for s in self.series:
            if len(s.values) != len(self.categories):
                raise ValueError(
                    f"계열 '{s.name}'의 값 개수({len(s.values)})가 categories({len(self.categories)})와 다릅니다.")
        if self.chart_type in {"pie", "doughnut"} and len(self.series) > 1:
            raise ValueError("pie/doughnut 차트는 계열을 1개만 지원합니다.")
        return self


class ImageSlide(_Slide):
    """이미지 슬라이드. 설명 불릿을 옆에 붙일 수 있다."""

    type: Literal["image"] = "image"
    title: str | None = Field(None, description="슬라이드 제목. full_bleed면 생략 가능.")
    image: ImageSource = Field(description="넣을 이미지.")
    caption: str | None = Field(None, description="이미지 아래 설명/출처.")
    placement: Literal["full", "left", "right", "full_bleed"] = Field(
        "full", description="full=본문 전체, left/right=이미지 위치(반대편에 bullets), full_bleed=여백 없이 꽉 채움.")
    bullets: list[BulletLike] = Field(
        default_factory=list, description="placement가 left/right일 때 반대편에 넣을 불릿.")


class QuoteSlide(_Slide):
    """인용문."""

    type: Literal["quote"] = "quote"
    text: str = Field(description="인용할 문장.")
    attribution: str | None = Field(None, description="말한 사람/출처.")
    role: str | None = Field(None, description="소속·직함 등 보조 정보.")


class KpiSlide(_Slide):
    """숫자 지표 카드 나열."""

    type: Literal["kpi"] = "kpi"
    title: str = Field(description="슬라이드 제목.")
    subtitle: str | None = Field(None, description="제목 아래 한 줄 설명.")
    items: list[KpiItem] = Field(description="지표 카드. 2~4개일 때 가장 보기 좋다.", min_length=1, max_length=6)


class TimelineSlide(_Slide):
    """가로 타임라인/단계 구성."""

    type: Literal["timeline"] = "timeline"
    title: str = Field(description="슬라이드 제목.")
    steps: list[TimelineStep] = Field(description="단계 목록. 3~5개 권장.", min_length=2, max_length=6)


class BlankSlide(_Slide):
    """정해진 레이아웃이 맞지 않을 때 쓰는 자유 배치 슬라이드."""

    type: Literal["blank"] = "blank"
    title: str | None = Field(None, description="제목. 생략하면 제목 없이 빈 슬라이드.")
    blocks: list[TextBlock] = Field(default_factory=list, description="직접 배치할 텍스트 상자들.")


AnySlide = Annotated[
    Union[
        TitleSlide, AgendaSlide, SectionSlide, BulletsSlide, TwoColumnSlide,
        ComparisonSlide, TableSlide, ChartSlide, ImageSlide, QuoteSlide,
        KpiSlide, TimelineSlide, BlankSlide,
    ],
    Field(discriminator="type"),
]

SLIDE_TYPES: tuple[str, ...] = (
    "title", "agenda", "section", "bullets", "two_column", "comparison",
    "table", "chart", "image", "quote", "kpi", "timeline", "blank",
)


# --- 덱 --------------------------------------------------------------------

class DeckSpec(_Base):
    """덱 한 벌 전체."""

    title: str = Field(description="덱 제목. 파일 속성과 표지 기본값으로 쓰인다.")
    subtitle: str | None = Field(None, description="덱 부제.")
    author: str | None = Field(None, description="작성자. 파일 속성에 기록된다.")
    slides: list[AnySlide] = Field(description="슬라이드 목록. 순서대로 생성된다.")

    theme: str | None = Field(
        None, description="테마 이름. list_themes로 확인. 생략하면 서버 기본 테마.")
    theme_overrides: dict[str, str] | None = Field(
        None, description="테마 색 일부 교체. 예: {'accent': '#0F62FE'}.")
    font_latin: str | None = Field(None, description="영문 폰트 이름 강제 지정.")
    font_ea: str | None = Field(None, description="한글(동아시아) 폰트 이름 강제 지정. 예: '맑은 고딕'.")

    template: str | None = Field(
        None, description="기반으로 쓸 .potx/.pptx 경로. 지정하면 그 파일의 레이아웃·마스터를 사용한다.")
    footer: str | None = Field(None, description="모든 본문 슬라이드 하단에 넣을 문구.")
    page_numbers: bool = Field(True, description="하단 오른쪽 페이지 번호 표시 여부.")
    aspect: Literal["16:9", "4:3"] = Field("16:9", description="슬라이드 비율. 템플릿 지정 시 무시된다.")


__all__ = [
    "DeckSpec", "AnySlide", "SLIDE_TYPES", "Bullet", "BulletLike", "Column",
    "ImageSource", "Series", "KpiItem", "TimelineStep", "TextBlock",
    "TitleSlide", "AgendaSlide", "SectionSlide", "BulletsSlide", "TwoColumnSlide",
    "ComparisonSlide", "TableSlide", "ChartSlide", "ImageSlide", "QuoteSlide",
    "KpiSlide", "TimelineSlide", "BlankSlide",
]

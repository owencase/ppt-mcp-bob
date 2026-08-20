"""덱 스펙.

tool 의 입력 스키마이자, LLM 이 보는 계약입니다.

두 가지를 의식하고 설계했습니다.

1. **kind 로 갈라지는 판별 유니온.** 슬라이드 종류마다 필요한 필드가 다릅니다.
   전부 Optional 로 한 덩어리에 몰아넣으면 모델이 어떤 조합이 유효한지 알 수
   없습니다. kind 를 리터럴로 나누면 JSON Schema 에 그대로 드러납니다.

2. **에러 메시지가 고치는 법을 말해준다.** 모델은 에러를 읽고 다시 시도합니다.
   "invalid input" 은 같은 실패를 반복하게 만들고, "불릿은 최대 6개입니다.
   7개를 보냈으니 슬라이드를 나누세요" 는 다음 시도를 성공시킵니다.
"""
from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, field_validator

from .theme import DEFAULT_THEME, THEMES

MAX_BULLETS = 6      # 한 장에 7개가 넘어가면 청중이 못 읽습니다
MAX_SLIDES = 40


class Strict(BaseModel):
    """모르는 필드를 거부하는 베이스.

    pydantic 기본값은 모르는 필드를 조용히 무시합니다. LLM 을 상대하는
    스키마에서는 그게 최악입니다 — 모델이 kind='title' 에 points 를 실어
    보내면 그 내용은 사라지는데 호출은 성공합니다. 모델은 자기가 뭘 잘못했는지
    영원히 모릅니다. 거부해야 다음 시도에 고칩니다.
    """
    model_config = ConfigDict(extra="forbid")


class TitleSlide(Strict):
    """표지."""
    kind: Literal["title"]
    title: str = Field(min_length=1, max_length=80)
    subtitle: str = Field("", max_length=120)


class SectionSlide(Strict):
    """구분 슬라이드. 장이 바뀔 때 씁니다."""
    kind: Literal["section"]
    title: str = Field(min_length=1, max_length=60)


class BulletsSlide(Strict):
    """제목 + 불릿. 가장 많이 쓰는 형태입니다."""
    kind: Literal["bullets"]
    title: str = Field(min_length=1, max_length=80)
    points: list[str] = Field(min_length=1)

    @field_validator("points")
    @classmethod
    def _check_points(cls, points: list[str]) -> list[str]:
        if len(points) > MAX_BULLETS:
            raise ValueError(
                f"불릿은 한 장에 최대 {MAX_BULLETS}개입니다 ({len(points)}개를 보냈습니다). "
                f"슬라이드를 나누세요."
            )
        for index, point in enumerate(points, 1):
            if not point.strip():
                raise ValueError(f"{index}번째 불릿이 비어 있습니다.")
            if len(point) > 120:
                raise ValueError(
                    f"{index}번째 불릿이 {len(point)}자입니다. 120자 이하로 줄이세요. "
                    f"불릿은 문장이 아니라 요점입니다."
                )
        return points


class ChartSlide(Strict):
    """제목 + 막대 차트. 숫자는 표보다 그림이 낫습니다."""
    kind: Literal["chart"]
    title: str = Field(min_length=1, max_length=80)
    series: dict[str, float] = Field(min_length=1)

    @field_validator("series")
    @classmethod
    def _check_series(cls, series: dict[str, float]) -> dict[str, float]:
        if len(series) > 12:
            raise ValueError(
                f"막대 12개가 넘으면 읽기 어렵습니다 ({len(series)}개). "
                f"상위 항목만 남기거나 묶으세요."
            )
        return series


Slide = Annotated[
    Union[TitleSlide, SectionSlide, BulletsSlide, ChartSlide],
    Field(discriminator="kind"),
]


class DeckSpec(Strict):
    """덱 하나 전체. 이걸 통째로 받아서 한 번에 만듭니다."""
    slides: list[Slide] = Field(min_length=1)
    theme: str = DEFAULT_THEME

    @field_validator("theme")
    @classmethod
    def _check_theme(cls, name: str) -> str:
        if name not in THEMES:
            raise ValueError(
                f"'{name}' 테마는 없습니다. 사용 가능: {', '.join(sorted(THEMES))}"
            )
        return name

    @field_validator("slides")
    @classmethod
    def _check_slides(cls, slides: list) -> list:
        if len(slides) > MAX_SLIDES:
            raise ValueError(f"슬라이드는 최대 {MAX_SLIDES}장입니다 ({len(slides)}장).")
        return slides

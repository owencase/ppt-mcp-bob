"""스펙 검증 테스트.

에러 메시지가 '고치는 법'을 말해주는지까지 검사합니다. 모델은 에러를 읽고
다시 시도하므로, 메시지 품질이 곧 다음 시도의 성공률입니다.
"""
import pytest
from pydantic import ValidationError

from ppt_mcp.models import DeckSpec


def _deck(**slide) -> dict:
    return {"theme": "tech_blue", "slides": [slide]}


class TestDiscriminatedUnion:
    def test_kind_selects_the_right_shape(self):
        spec = DeckSpec.model_validate(_deck(kind="bullets", title="요약", points=["가", "나"]))
        assert spec.slides[0].kind == "bullets"

    def test_wrong_field_for_kind_is_rejected(self):
        # title 슬라이드에 points 를 주는 건 잘못입니다. 판별 유니온이 잡아냅니다.
        with pytest.raises(ValidationError):
            DeckSpec.model_validate(_deck(kind="title", title="제목", points=["가"]))

    def test_unknown_kind_is_rejected(self):
        with pytest.raises(ValidationError):
            DeckSpec.model_validate(_deck(kind="표", title="제목"))


class TestErrorMessagesTeach:
    def test_too_many_bullets_says_to_split(self):
        with pytest.raises(ValidationError) as exc:
            DeckSpec.model_validate(_deck(kind="bullets", title="많음",
                                          points=[f"항목{i}" for i in range(9)]))
        assert "슬라이드를 나누세요" in str(exc.value)

    def test_long_bullet_explains_why(self):
        with pytest.raises(ValidationError) as exc:
            DeckSpec.model_validate(_deck(kind="bullets", title="긺", points=["가" * 200]))
        assert "요점입니다" in str(exc.value)

    def test_unknown_theme_lists_the_valid_ones(self):
        with pytest.raises(ValidationError) as exc:
            DeckSpec.model_validate({"theme": "없는테마",
                                     "slides": [{"kind": "section", "title": "장"}]})
        assert "tech_blue" in str(exc.value)

    def test_empty_bullet_names_the_position(self):
        with pytest.raises(ValidationError) as exc:
            DeckSpec.model_validate(_deck(kind="bullets", title="빔", points=["가", "  ", "다"]))
        assert "2번째" in str(exc.value)

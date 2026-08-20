"""레이아웃 테스트.

레이아웃은 python-pptx 를 모르는 순수 함수라 파일을 만들지 않고 검사할 수
있습니다. 그래서 빠르고, 규칙을 촘촘히 확인할 수 있습니다. 레이아웃을 렌더링
코드에 섞어 놓으면 이런 테스트를 못 씁니다 — 그게 이 모듈을 분리한 이유입니다.
"""
import pytest

from ppt_mcp import layout


class TestGeometry:
    # 모든 슬라이드 종류가 슬라이드 밖으로 나가지 않아야 합니다.
    # 디자인을 고치다 좌표를 잘못 넣으면 여기서 바로 걸립니다.
    def test_all_layouts_stay_inside_the_slide(self):
        candidates = [layout.title_slide(), layout.section_slide(),
                      layout.chart_slide(), layout.bullets_slide(4)]
        for boxes in candidates:
            for name, box in boxes.items():
                assert layout.within_slide(box), f"{name} 이 슬라이드를 벗어납니다: {box}"

    def test_bullet_slots_do_not_overlap(self):
        boxes = layout.bullets_slide(5)
        slot = boxes["bullet_slot"]
        # 슬롯을 순서대로 쌓았을 때 본문 영역을 넘지 않아야 합니다.
        assert slot.top + 5 * slot.height <= boxes["body"].bottom + 0.01

    def test_title_never_collides_with_body(self):
        boxes = layout.bullets_slide(3)
        assert boxes["title"].bottom <= boxes["body"].top


class TestTextFitting:
    def test_hangul_is_wider_than_latin(self):
        # 한글은 라틴 문자보다 넓습니다. 같게 계산하면 과소추정해서 넘칩니다.
        latin = layout.estimate_lines("a" * 40, 20, 10)
        hangul = layout.estimate_lines("가" * 40, 20, 10)
        assert hangul > latin

    def test_newlines_start_new_lines(self):
        assert layout.estimate_lines("a\nb\nc", 20, 30) == 3

    def test_font_shrinks_to_fit(self):
        box = layout.Box(0, 0, 28.8, 4.2)          # 실제 표지 제목 상자 크기
        short = layout.fit_font_size("짧은 제목", box, 44, 20)
        long_ = layout.fit_font_size("아주 " * 20 + "긴 제목", box, 44, 14)
        assert short == 44                # 짧으면 요청한 크기 그대로
        assert long_ < short              # 길면 줄어듦

    def test_impossible_text_fails_loudly(self):
        # 조용히 잘리게 두면 인쇄하고 나서야 발견됩니다. 실패시키는 게 낫습니다.
        box = layout.Box(0, 0, 5, 1)
        with pytest.raises(ValueError, match="줄여도"):
            layout.fit_font_size("가" * 2000, box, 44, layout.MIN_BODY_PT)

    def test_failure_message_tells_you_what_to_do(self):
        box = layout.Box(0, 0, 5, 1)
        with pytest.raises(ValueError) as exc:
            layout.fit_font_size("가" * 2000, box, 44, layout.MIN_BODY_PT)
        assert "슬라이드를 나누세요" in str(exc.value)

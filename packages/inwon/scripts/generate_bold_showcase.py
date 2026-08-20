from __future__ import annotations

import os
from pathlib import Path

from canva_ppt_mcp.models import ContentItem, DeckPlan, DesignSystem, Palette, SlideSpec, Typography
from canva_ppt_mcp.pipeline import create_presentation


def build(output: Path) -> None:
    os.environ.pop("OPENAI_API_KEY", None)
    design = DesignSystem(
        palette=Palette(
            primary="#111A3A",
            secondary=["#7357FF", "#193B73"],
            accent="#00F0C8",
            background_light="#F4F7FF",
            background_dark="#070B1E",
        ),
        typography=Typography(
            header_font="Bookman Old Style",
            body_font="Arial",
            title_size=56,
            section_header_size=24,
            body_size=16,
            caption_size=11,
        ),
        visual_motif="전기적 청록색 광원과 비대칭 레이어",
        style_preset="neon",
        layout_rotation=["title", "big_stat", "two_column", "image_focus", "grid_2x2", "comparison", "timeline"],
        visual_intensity="maximal",
        dark_slide_ratio=.50,
        gradient_backgrounds=True,
        background_texture=True,
        dynamic_composition=True,
    )
    plan = DeckPlan(
        communication_job="AI 제품 전략의 핵심과 실행 우선순위를 강한 시각 위계로 전달한다.",
        design_system=design,
        slides=[
            SlideSpec(
                title="AI는 기능이 아니라 새로운 인터페이스다",
                subtitle="제품 경험을 다시 설계하는 2027 전략",
                layout="title",
            ),
            SlideSpec(
                title="첫 경험이 채택 속도를 결정한다",
                layout="big_stat",
                items=[ContentItem(
                    heading="초기 가치 인지",
                    body="사용자는 복잡한 기능 목록보다 첫 상호작용에서 얻는 명확한 결과로 제품을 판단합니다.",
                    value="10s",
                )],
            ),
            SlideSpec(
                title="기술보다 경험의 연결성이 중요하다",
                layout="two_column",
                items=[
                    ContentItem(heading="분절된 기능", body="도구마다 맥락이 끊기면 사용자는 매번 다시 설명해야 합니다."),
                    ContentItem(heading="연결된 경험", body="의도와 작업 흐름이 이어질 때 AI는 하나의 인터페이스가 됩니다."),
                ],
            ),
            SlideSpec(
                title="하나의 강한 순간이 제품을 기억하게 한다",
                layout="image_focus",
                items=[ContentItem(
                    heading="Hero interaction",
                    body="제품을 대표하는 한 가지 상호작용을 크게 보여주고 나머지는 보조 정보로 낮춥니다.",
                )],
            ),
            SlideSpec(
                title="네 가지 원칙이 경험을 완성한다",
                layout="grid_2x2",
                items=[
                    ContentItem(heading="맥락", body="사용자의 목표와 현재 상태를 이해합니다."),
                    ContentItem(heading="속도", body="첫 유용한 결과까지의 단계를 줄입니다."),
                    ContentItem(heading="신뢰", body="근거와 수정 가능성을 함께 제공합니다."),
                    ContentItem(heading="연결", body="다음 행동으로 자연스럽게 이어집니다."),
                ],
            ),
            SlideSpec(
                title="기능 중심에서 결과 중심으로 전환한다",
                layout="comparison",
                items=[
                    ContentItem(heading="Before", body="기능 수와 모델 사양을 중심으로 제품을 설명합니다."),
                    ContentItem(heading="After", body="사용자가 더 빨리 끝내는 일과 얻는 결과를 보여줍니다."),
                ],
            ),
            SlideSpec(
                title="작게 검증하고 빠르게 확장한다",
                layout="timeline",
                items=[
                    ContentItem(heading="발견", body="가장 반복적인 사용자 문제를 고릅니다."),
                    ContentItem(heading="검증", body="한 가지 핵심 상호작용을 시험합니다."),
                    ContentItem(heading="통합", body="실제 작업 흐름과 연결합니다."),
                    ContentItem(heading="확장", body="측정된 가치가 있는 영역부터 넓힙니다."),
                ],
            ),
            SlideSpec(
                title="다음 제품은 설명보다 경험으로 증명해야 한다",
                subtitle="첫 번째 Hero interaction과 성공 기준을 이번 주에 확정하세요.",
                layout="closing",
                items=[ContentItem(heading="NEXT", body="한 가지 사용자 문제와 책임자를 선택합니다.")],
            ),
        ],
    )
    create_presentation(
        topic="AI 제품 전략",
        output_path=str(output),
        slide_count=8,
        language="ko",
        content_json=plan.model_dump(),
        max_qa_rounds=3,
    )


if __name__ == "__main__":
    build(Path("output/bold-visual-showcase.pptx"))

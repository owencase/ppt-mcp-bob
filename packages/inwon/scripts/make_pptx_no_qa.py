"""
PPTX 생성 스크립트 (QA 렌더링 없이)
Usage: python scripts/make_pptx_no_qa.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from canva_ppt_mcp.planner import create_plan
from canva_ppt_mcp.render import render_auto_deck, write_speaker_notes

topic = "세명컴퓨터고등학교 학교 소개"
audience = "학생, 학부모, 입학 희망자"
purpose = "세명컴퓨터고등학교의 교육 목표, 학과 구성, 특색 교육, 시설, 졸업 후 진로 소개"
slide_count = 10
language = "ko"
style_preference = "organic"
output_path = "output/semyung_computer_hs.pptx"

print(f"[1/3] 슬라이드 계획 생성 중... (주제: {topic})")
plan = create_plan(
    topic=topic,
    audience=audience,
    purpose=purpose,
    slide_count=slide_count,
    language=language,
    content_json=None,
    research_text=None,
    source_urls=None,
    research_required=True,
    style_preference=style_preference,
    research_documents=None,
)

print(f"[2/3] PPTX 렌더링 중... -> {output_path}")
Path(output_path).parent.mkdir(parents=True, exist_ok=True)
render_auto_deck(plan, output_path)
write_speaker_notes(output_path, plan)

print(f"[3/3] 완료: {Path(output_path).resolve()}")
print(f"  슬라이드 수: {len(plan.slides)}")
print(f"  스타일: {plan.design_system.style_preset}")
print(f"  팔레트 primary: {plan.design_system.palette.primary}")
for i, slide in enumerate(plan.slides, 1):
    print(f"  [{i:02d}] [{slide.layout:12s}] {slide.title}")

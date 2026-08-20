"""브릿지를 직접 호출하는 최소 예제.

MCP 서버 없이 bridge.py 만으로 덱을 만듭니다. 프로토콜이 어떻게 생겼는지
한눈에 보려는 용도입니다.

    cd reference/ppt-bridge
    .venv/bin/python example_deck.py

산출물 example.pptx 는 현재 디렉터리에 생기고 .gitignore 대상입니다.
"""
import json
import subprocess
import sys
from pathlib import Path

BRIDGE = Path(__file__).with_name("bridge.py")
DECK = "example.pptx"


def call(action: str, **params) -> dict:
    """stdin 으로 JSON 한 줄 넣고 stdout 으로 JSON 한 줄 받습니다. 그게 전부입니다."""
    request = json.dumps({"action": action, "params": {"file_path": DECK, **params}})
    done = subprocess.run([sys.executable, str(BRIDGE)], input=request,
                          capture_output=True, encoding="utf-8")
    result = json.loads(done.stdout)
    if not result.get("success"):
        raise RuntimeError(f"{action} 실패: {result.get('error')}")
    return result


# 1) 빈 프레젠테이션 → 2) 슬라이드 → 3) 내용 → 4) 테마
call("create_presentation")
call("add_slide")

# 좌표와 크기는 cm, 색은 '#' 없는 RRGGBB 입니다.
# 이 값들을 호출자가 직접 계산해야 한다는 점에 주목하세요 — 설계상의 선택이고,
# 장단점은 MCP-DESIGN.md 1번 축(tool 고도)에 정리돼 있습니다.
call("add_text_box", slide_index=0, text="브릿지 직접 호출 예제",
     left_cm=2.5, top_cm=6.5, width_cm=28.8, height_cm=3.0,
     font_size_pt=40, bold=True, color_hex="E2E8F0", align="center")

theme = call("apply_theme", theme="tech_blue")["data"]["theme"]
print(f"테마 적용: bg #{theme['bg']}, accent #{theme['accent1']}")

info = call("get_presentation_info")["data"]
print(f"완성: {DECK} — {info['slide_count']}장, {info['width_cm']}×{info['height_cm']}cm")

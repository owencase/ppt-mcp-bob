"""bridge.py 테스트.

실행:
    cd packages/ppt-bridge
    .venv/bin/python -m pytest tests -v

테스트를 왜 이렇게 나눠 놨는지 각 클래스 주석에 적어 뒀습니다.
새 액션을 추가할 때 이 파일을 흉내 내서 테스트도 같이 추가해 주세요.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

import bridge

PACKAGE_ROOT = Path(__file__).resolve().parent.parent
BRIDGE_PY = PACKAGE_ROOT / "bridge.py"


def call(action: str, **params) -> dict:
    """dispatch() 를 직접 호출합니다. 프로세스를 안 띄우니 빠릅니다."""
    return bridge.dispatch(json.dumps({"action": action, "params": params}))


# ───────────────────────────────────────────────────────────────────────────
# 1. 순수 함수
#
# 단위 변환과 색 변환은 모든 액션이 공유합니다. 여기가 틀리면 전부 틀립니다.
# ───────────────────────────────────────────────────────────────────────────
class TestPureFunctions:
    def test_cm_to_emu(self):
        assert bridge.cm(1) == 360_000
        assert bridge.cm(0) == 0
        assert bridge.cm(2.5) == 900_000

    def test_hex_to_rgb(self):
        assert tuple(bridge.hex_to_rgb("FF8000")) == (255, 128, 0)

    def test_hex_to_rgb_accepts_leading_hash(self):
        # 레포 규약은 '#' 없는 RRGGBB 지만, 붙여서 보내도 받아 줍니다.
        assert tuple(bridge.hex_to_rgb("#FF8000")) == (255, 128, 0)


# ───────────────────────────────────────────────────────────────────────────
# 2. 프로토콜 계약
#
# mcp-server 는 stdout 한 줄을 JSON.parse 합니다. 어떤 입력이 와도 파싱
# 가능한 JSON 이 나와야 하고, 실패는 예외가 아니라 success:false 여야 합니다.
# 이게 깨지면 LLM 클라이언트에는 "Invalid JSON from bridge" 라는 아무 정보
# 없는 문자열만 보입니다.
# ───────────────────────────────────────────────────────────────────────────
class TestProtocol:
    def test_unknown_action_returns_error_not_exception(self):
        result = call("이런_액션_없음")
        assert result["success"] is False
        assert "이런_액션_없음" in result["error"]

    def test_unknown_action_lists_available_actions(self):
        # 사람이 읽는 메시지 말고, 기계가 읽을 수 있는 목록도 같이 줍니다.
        # mcp-server 쪽 테스트가 이 목록으로 tool 정의와 대조합니다.
        result = call("없음")
        assert result["data"]["available"] == sorted(bridge.HANDLERS)

    def test_broken_json_returns_error(self):
        result = bridge.dispatch("{이건 JSON 이 아닙니다")
        assert result["success"] is False
        assert "Invalid JSON" in result["error"]

    def test_non_object_request_returns_error(self):
        assert bridge.dispatch("[1, 2, 3]")["success"] is False

    def test_missing_required_param_names_the_param(self):
        # file_path 없이 호출. 스택트레이스 대신 빠진 이름을 알려줘야 합니다.
        result = call("add_text_box", slide_index=0)
        assert result["success"] is False
        assert "file_path" in result["error"]


# ───────────────────────────────────────────────────────────────────────────
# 3. stdout 오염 방어
#
# 디버깅하다 print() 를 남기는 일은 반드시 일어납니다. 그때 프로토콜이
# 깨지면 원인을 찾기가 매우 어렵습니다. 그래서 stdout 으로 나간 출력은
# stderr 로 옮기고, 응답 JSON 은 멀쩡히 나오도록 해 뒀습니다.
#
# 이 테스트는 진짜 프로세스를 띄웁니다. mcp-server 가 하는 것과 똑같이요.
# ───────────────────────────────────────────────────────────────────────────
class TestStdoutIsProtocolOnly:
    NOISY_HANDLER = (
        "import sys; sys.path.insert(0, {root!r})\n"
        "import bridge\n"
        "def noisy(params):\n"
        "    print('디버깅용 출력입니다')\n"
        "    print('두 번째 줄')\n"
        "    return bridge.ok('끝')\n"
        "bridge.HANDLERS['noisy'] = noisy\n"
        "bridge.main()\n"
    )

    def test_stray_print_does_not_corrupt_response(self, tmp_path):
        script = tmp_path / "noisy_bridge.py"
        script.write_text(self.NOISY_HANDLER.format(root=str(PACKAGE_ROOT)), encoding="utf-8")

        proc = subprocess.run(
            [sys.executable, str(script)],
            input='{"action": "noisy", "params": {}}',
            capture_output=True,
            encoding="utf-8",
        )

        # stdout 은 여전히 파싱 가능한 JSON 한 줄
        response = json.loads(proc.stdout)
        assert response["success"] is True

        # 출력은 삼켜지지 않고 stderr 에서 볼 수 있어야 합니다
        assert "디버깅용 출력입니다" in proc.stderr
        assert "두 번째 줄" in proc.stderr

    def test_real_process_roundtrip(self, tmp_path):
        # 실제 프로세스로 한 번은 왕복해 봅니다. main() 의 인코딩 처리까지 포함.
        target = tmp_path / "한글파일.pptx"
        proc = subprocess.run(
            [sys.executable, str(BRIDGE_PY)],
            input=json.dumps({"action": "create_presentation",
                              "params": {"file_path": str(target)}}),
            capture_output=True,
            encoding="utf-8",
        )
        assert proc.returncode == 0
        assert json.loads(proc.stdout)["success"] is True
        assert target.exists()


# ───────────────────────────────────────────────────────────────────────────
# 4. 액션 동작
#
# 실제로 .pptx 가 만들어지고 읽히는지. tmp_path 는 pytest 가 테스트마다
# 새로 만들어 주는 임시 디렉터리라 산출물이 레포에 남지 않습니다.
# ───────────────────────────────────────────────────────────────────────────
class TestActions:
    @pytest.fixture
    def deck(self, tmp_path) -> str:
        path = str(tmp_path / "test.pptx")
        assert call("create_presentation", file_path=path)["success"]
        return path

    def test_create_uses_widescreen_by_default(self, deck):
        info = call("get_presentation_info", file_path=deck)
        assert info["success"] is True
        assert info["data"]["width_cm"] == pytest.approx(33.87, abs=0.01)
        assert info["data"]["slide_count"] == 0

    def test_add_slide_returns_index_of_new_slide(self, deck):
        first = call("add_slide", file_path=deck)
        second = call("add_slide", file_path=deck)
        assert first["data"]["slide_index"] == 0
        assert second["data"]["slide_index"] == 1

    def test_text_box_lands_on_the_slide(self, deck):
        call("add_slide", file_path=deck)
        assert call("add_text_box", file_path=deck, slide_index=0, text="안녕하세요",
                    left_cm=2, top_cm=2, width_cm=10, height_cm=3)["success"]

        info = call("get_presentation_info", file_path=deck)
        shapes = info["data"]["slides"][0]["shapes"]
        assert len(shapes) == 1
        assert shapes[0]["left_cm"] == pytest.approx(2.0)

    def test_apply_theme_returns_palette(self, deck):
        call("add_slide", file_path=deck)
        result = call("apply_theme", file_path=deck, theme="tech_blue")
        assert result["success"] is True
        # 배경만 칠하고 나머지 색은 응답으로 돌려주는 계약입니다.
        # 호출자가 이 값을 다음 add_text_box 에 넣어 씁니다.
        assert result["data"]["theme"]["accent1"] == "38BDF8"

    def test_unknown_theme_is_rejected(self, deck):
        result = call("apply_theme", file_path=deck, theme="없는테마")
        assert result["success"] is False
        assert "minimal_dark" in result["error"]   # 쓸 수 있는 목록을 알려줌

"""MCP 계약 테스트.

서버가 부팅 중에 죽으면 클라이언트에는 그냥 "tool 이 없음" 으로 보입니다.
원인을 찾기 어려운 실패라 여기서 잡습니다.

tool 뿐 아니라 resource 와 prompt 도 확인합니다. 셋 다 쓰는 게 이 서버의
설계 의도인데, 하나가 조용히 빠져도 아무도 모르면 의도가 무너집니다.

MCP 목록 API 는 async 지만 pytest-asyncio 를 쓰지 않습니다. asyncio.run 으로
감싸면 되는 일에 의존성을 하나 더 늘릴 이유가 없습니다.
"""
import asyncio

import pytest

from ppt_mcp.server import mcp


def run(coro):
    return asyncio.run(coro)


@pytest.fixture(scope="module")
def tools():
    return run(mcp.list_tools())


class TestTools:
    def test_exactly_three_tools(self, tools):
        # tool 이 적을수록 모델이 덜 헷갈립니다. 늘어나면 의식적으로 늘리세요.
        assert {tool.name for tool in tools} == {"create_deck", "validate_deck", "describe_deck"}

    def test_every_tool_has_a_real_description(self, tools):
        # description 은 모델이 이 tool 을 언제 쓸지 판단하는 유일한 근거입니다.
        for tool in tools:
            assert tool.description and len(tool.description) > 30, tool.name

    def test_create_deck_takes_a_whole_spec_not_coordinates(self, tools):
        """이 서버의 설계 주장을 고정하는 테스트.

        tool 인자에 좌표나 폰트 크기가 등장하면 레이아웃 책임이 호출자에게
        새어나간 것입니다. 그 순간 이 서버는 레퍼런스 구현과 같아집니다.
        """
        tool = next(t for t in tools if t.name == "create_deck")
        # SDK 2.x 는 input_schema, 1.x 는 inputSchema 로 노출합니다.
        schema = getattr(tool, "input_schema", None) or tool.inputSchema
        assert set(schema["properties"]) == {"spec", "output_path"}

        leaked = [word for word in ("left_cm", "top_cm", "width_cm", "height_cm",
                                    "font_size_pt", "color_hex", "align")
                  if word in str(schema)]
        assert leaked == [], f"레이아웃 인자가 tool 표면에 노출됨: {leaked}"


class TestResourcesAndPrompts:
    def test_theme_list_is_a_resource_not_a_tool(self):
        # 테마는 '읽는' 데이터입니다. tool 로 만들면 색을 알아내려고 실행해야 합니다.
        uris = {str(resource.uri) for resource in run(mcp.list_resources())}
        assert "theme://list" in uris

    def test_templated_resource_is_exposed(self):
        """theme://{name} 은 목록이 아니라 템플릿으로 등록됩니다.

        MCP 는 고정 URI 와 패턴 URI 를 나눠서 광고합니다. list_resources 만
        확인하면 템플릿이 통째로 빠져도 모릅니다.
        """
        templates = run(mcp.list_resource_templates())
        patterns = {getattr(t, "uriTemplate", None) or t.uri_template for t in templates}
        assert "theme://{name}" in patterns

    def test_theme_resource_returns_the_full_palette(self):
        # 모델이 색을 지어내지 않으려면 여기서 실제 값을 받아야 합니다.
        body = run(mcp.read_resource("theme://list"))
        text = body[0].content if isinstance(body, list) else str(body)
        for key in ("bg", "text", "accent", "muted", "tech_blue"):
            assert key in text

    def test_prompt_is_exposed(self):
        names = {prompt.name for prompt in run(mcp.list_prompts())}
        assert "deck_from_topic" in names

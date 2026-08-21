"""실행 진입점. stdio가 기본이고, 사내 배포용으로 Streamable HTTP도 지원한다.

  python -m ppt_mcp                                   # stdio (BOB 로컬 연결)
  python -m ppt_mcp --transport http --port 8080      # 원격 배포
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from dataclasses import replace
from pathlib import Path

from .config import Settings
from .server import build_server
from .theme import list_theme_names


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="ppt-mcp", description="PPT(.pptx) 생성·편집 MCP 서버")
    parser.add_argument("--transport", choices=("stdio", "http"), default="stdio",
                        help="연결 방식. 기본값 stdio.")
    parser.add_argument("--host", default="127.0.0.1", help="http 모드에서 바인딩할 호스트.")
    parser.add_argument("--port", type=int, default=8080, help="http 모드 포트.")
    parser.add_argument("--path", default="/mcp", help="http 모드 엔드포인트 경로.")
    parser.add_argument("--output-dir", help="생성한 파일을 저장할 디렉터리.")
    parser.add_argument("--template-dir", help="템플릿을 찾을 디렉터리.")
    parser.add_argument("--default-template", help="template을 지정하지 않았을 때 쓸 기본 템플릿.")
    parser.add_argument("--theme", choices=list_theme_names(), help="기본 테마.")
    parser.add_argument("--allow-remote-images", action="store_true",
                        help="이미지 URL 다운로드를 허용한다(기본 차단).")
    return parser.parse_args(argv)


def _settings_from(args: argparse.Namespace) -> Settings:
    settings = Settings.from_env()
    overrides: dict = {}
    if args.output_dir:
        overrides["output_dir"] = Path(args.output_dir).expanduser().resolve()
    if args.template_dir:
        overrides["template_dir"] = Path(args.template_dir).expanduser()
    if args.default_template:
        overrides["default_template"] = Path(args.default_template).expanduser()
    if args.theme:
        overrides["default_theme"] = args.theme
    if args.allow_remote_images:
        overrides["allow_remote_images"] = True
    return replace(settings, **overrides) if overrides else settings


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    settings = _settings_from(args)
    settings.ensure_output_dir()
    mcp = build_server(settings)

    if args.transport == "stdio":
        # stdout은 MCP 프로토콜 전용이므로 로그는 stderr로만 낸다.
        print(f"[ppt-mcp] stdio 시작 · 저장 위치: {settings.output_dir}", file=sys.stderr)
        mcp.run("stdio")
    else:
        print(f"[ppt-mcp] http 시작 · http://{args.host}:{args.port}{args.path} "
              f"· 저장 위치: {settings.output_dir}", file=sys.stderr)
        asyncio.run(mcp.run_streamable_http_async(
            host=args.host, port=args.port, streamable_http_path=args.path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

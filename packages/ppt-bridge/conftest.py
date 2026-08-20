"""pytest 설정.

tests/ 안에서 `import bridge` 가 되도록 패키지 루트를 sys.path 에 넣어 줍니다.
이 파일이 없으면 `pytest` 를 어디서 실행하느냐에 따라 import 가 되기도 하고
안 되기도 합니다.
"""
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))

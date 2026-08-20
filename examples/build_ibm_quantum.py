# -*- coding: utf-8 -*-
"""
IBM 양자컴퓨터 PPT 5장 빌드 스크립트 (ppt-bridge 사용 예시)

레포 루트에서:
    python3 -X utf8 examples/build_ibm_quantum.py

결과물 IBM_Quantum.pptx 는 현재 작업 디렉터리에 생성되며, .gitignore 대상입니다.
"""
import sys, json, subprocess, os

REPO_ROOT   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PYTHON_BIN  = os.environ.get("PYTHON_BIN", sys.executable)
BRIDGE      = os.environ.get(
    "BRIDGE_SCRIPT",
    os.path.join(REPO_ROOT, "reference", "ppt-bridge", "bridge.py"),
)
FILE        = "IBM_Quantum.pptx"

# ── 색상 팔레트 (tech_blue 계열) ──────────────────────────────────────────
BG        = "0F172A"   # 배경 (딥 네이비)
WHITE     = "E2E8F0"   # 본문 텍스트
BLUE      = "38BDF8"   # 메인 액센트
CYAN      = "22D3EE"   # 보조 액센트
GRAY      = "94A3B8"   # 설명 텍스트
DIVIDER   = "1E3A5F"   # 구분선/박스
DARK_CARD = "0D1F35"   # 카드 배경


def bridge(action: str, **params):
    payload = json.dumps({"action": action, "params": {"file_path": FILE, **params}}, ensure_ascii=False)
    r = subprocess.run(
        [PYTHON_BIN, "-X", "utf8", BRIDGE],
        input=payload, capture_output=True, encoding="utf-8"
    )
    result = json.loads(r.stdout)
    status = "OK" if result.get("success") else f"ERR: {result.get('error')}"
    print(f"  [{action}] {status}")
    return result


def txt(si, text, l, t, w, h, size=18, bold=False, color=WHITE, align="left"):
    bridge("add_text_box", slide_index=si, text=text,
           left_cm=l, top_cm=t, width_cm=w, height_cm=h,
           font_size_pt=size, bold=bold, color_hex=color, align=align)


def rect(si, l, t, w, h, fill=BLUE, line=None):
    kw = dict(slide_index=si, shape_type="rectangle",
              left_cm=l, top_cm=t, width_cm=w, height_cm=h, fill_color_hex=fill)
    if line:
        kw["line_color_hex"] = line
    bridge("add_shape", **kw)


def rrect(si, l, t, w, h, fill=BLUE):
    bridge("add_shape", slide_index=si, shape_type="rounded_rectangle",
           left_cm=l, top_cm=t, width_cm=w, height_cm=h, fill_color_hex=fill)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 준비 — 빈 프레젠테이션 + 슬라이드 5장
# 아래 슬라이드들은 slide_index 0~4 에 그리므로 먼저 만들어 두어야 합니다.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SLIDE_COUNT = 5

print("[준비] 프레젠테이션 생성")
bridge("create_presentation")
for _ in range(SLIDE_COUNT):
    bridge("add_slide", layout_index=6)          # 6 = 빈 레이아웃
for si in range(SLIDE_COUNT):
    bridge("set_background_color", slide_index=si, color_hex=BG)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Slide 0 — 타이틀 슬라이드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
print("\n[Slide 0] 타이틀")
rect(0, 0, 0, 0.55, 19.05, fill=BLUE)          # 왼쪽 강조 바
rect(0, 0.55, 15.2, 33.32, 0.06, fill=DIVIDER) # 가로 구분선

txt(0, "IBM QUANTUM",          2.2,  2.8, 30, 1.4, size=14, bold=True,  color=BLUE)
txt(0, "양자컴퓨터의 현재와 미래",  2.2,  4.4, 30, 5.0, size=50, bold=True,  color=WHITE)
txt(0, "초전도 큐비트부터 Qiskit까지 — IBM이 이끄는 양자 혁명", 2.2, 9.8, 30, 2.0, size=20, color=GRAY)

# 오른쪽 장식 박스들
rrect(0, 26.0,  4.0, 6.5, 2.5, fill=DIVIDER)
rrect(0, 26.0,  7.2, 6.5, 2.5, fill=DIVIDER)
rrect(0, 26.0, 10.4, 6.5, 2.5, fill=DIVIDER)
txt(0, "127+ Qubits",  26.3,  4.3, 6.0, 1.0, size=16, bold=True,  color=BLUE,  align="center")
txt(0, "Eagle Processor",26.3,  5.1, 6.0, 0.8, size=11, color=GRAY, align="center")
txt(0, "Qiskit SDK",   26.3,  7.5, 6.0, 1.0, size=16, bold=True,  color=CYAN,  align="center")
txt(0, "오픈소스 양자 프레임워크", 26.3, 8.3, 6.0, 0.8, size=11, color=GRAY, align="center")
txt(0, "IBM Quantum Network", 26.3, 10.7, 6.0, 1.0, size=14, bold=True, color=WHITE, align="center")
txt(0, "180+ 글로벌 파트너", 26.3, 11.5, 6.0, 0.8, size=11, color=GRAY, align="center")

rect(0, 2.2, 15.5, 6.0, 0.06, fill=BLUE)
txt(0, "IBM Quantum Computing  |  2024",  2.2, 15.8, 20, 1.0, size=12, color="475569")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Slide 1 — 양자컴퓨터란 무엇인가?
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
print("\n[Slide 1] 양자컴퓨터란?")
rect(1, 0, 0, 33.87, 4.8, fill="071526")       # 상단 헤더 영역
txt(1, "01",                    2.0, 0.6, 4.0, 1.5, size=13, bold=True, color=BLUE)
txt(1, "양자컴퓨터란 무엇인가?", 2.0, 1.6, 28, 2.8, size=36, bold=True, color=WHITE)
rect(1, 2.0, 4.6, 29.87, 0.06, fill=BLUE)      # 구분선

# 왼쪽: 고전 컴퓨터 카드
rect(1, 2.0, 5.5, 13.5, 9.5, fill=DARK_CARD, line=DIVIDER)
txt(1, "고전 컴퓨터",     2.5,  5.9, 12.5, 1.2, size=18, bold=True, color=GRAY)
txt(1, "0 또는 1 (비트)",  2.5,  7.2, 12.5, 1.0, size=15, bold=True, color=WHITE)
txt(1, "순차적 연산 처리\n한 번에 하나의 상태\n확정적(결정론적) 계산\n기존 알고리즘 최적화", 2.5, 8.4, 12.5, 5.0, size=14, color=GRAY)

# 오른쪽: 양자 컴퓨터 카드
rect(1, 17.2, 5.5, 14.2, 9.5, fill=DARK_CARD, line=BLUE)
txt(1, "양자 컴퓨터",     17.7,  5.9, 13.0, 1.2, size=18, bold=True, color=BLUE)
txt(1, "0 과 1 동시 (큐비트)", 17.7, 7.2, 13.0, 1.0, size=15, bold=True, color=WHITE)
txt(1, "병렬 양자 연산\n중첩·얽힘 활용\n확률적 측정 결과\n지수적 연산 가속", 17.7, 8.4, 13.0, 5.0, size=14, color=CYAN)

# 중앙 VS 표시
rrect(1, 15.3, 9.2, 2.0, 1.8, fill="1E3A5F")
txt(1, "VS", 15.3, 9.4, 2.0, 1.4, size=18, bold=True, color=WHITE, align="center")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Slide 2 — IBM 양자 하드웨어
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
print("\n[Slide 2] IBM 양자 하드웨어")
rect(2, 0, 0, 33.87, 4.8, fill="071526")
txt(2, "02",                  2.0, 0.6, 4.0, 1.5, size=13, bold=True, color=BLUE)
txt(2, "IBM 양자 하드웨어 로드맵", 2.0, 1.6, 28, 2.8, size=36, bold=True, color=WHITE)
rect(2, 2.0, 4.6, 29.87, 0.06, fill=BLUE)

# 프로세서 로드맵 카드 3개
for i, (name, qubits, year, desc, col) in enumerate([
    ("Eagle",  "127 큐비트", "2021", "최초 100+큐비트\n오류율 대폭 감소\n헤비-헥스 위상",  "1E3A5F"),
    ("Osprey", "433 큐비트", "2022", "멀티칩 모듈\n개선된 연결성\n대규모 얽힘 가능", "0D2A4A"),
    ("Condor", "1,121 큐비트","2023", "세계 최초 1000+\n큐비트 프로세서\n양자 우위 실증 기반", "072040"),
]):
    lx = 2.0 + i * 10.7
    rrect(2, lx, 5.3, 9.8, 10.5, fill=col)
    rect(2,  lx, 5.3, 9.8, 0.35, fill=BLUE if i==2 else (CYAN if i==1 else "3B82F6"))
    txt(2, name,   lx+0.4, 5.8,  9.0, 1.2, size=22, bold=True, color=WHITE)
    txt(2, qubits, lx+0.4, 7.1,  9.0, 1.0, size=17, bold=True, color=BLUE if i==2 else CYAN)
    txt(2, year,   lx+0.4, 8.0,  9.0, 0.8, size=12, color=GRAY)
    rect(2, lx+0.4, 8.9, 8.8, 0.04, fill=DIVIDER)
    txt(2, desc,   lx+0.4, 9.2,  9.0, 4.5, size=13, color="CBD5E1")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Slide 3 — Qiskit & 실제 응용 분야
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
print("\n[Slide 3] Qiskit & 응용 분야")
rect(3, 0, 0, 33.87, 4.8, fill="071526")
txt(3, "03",                        2.0, 0.6, 4.0, 1.5, size=13, bold=True, color=BLUE)
txt(3, "Qiskit & 실제 응용 분야",   2.0, 1.6, 28, 2.8, size=36, bold=True, color=WHITE)
rect(3, 2.0, 4.6, 29.87, 0.06, fill=BLUE)

# 왼쪽: Qiskit 설명
rect(3, 2.0, 5.4, 14.5, 10.5, fill=DARK_CARD, line=DIVIDER)
txt(3, "Qiskit SDK",        2.5, 5.8, 13.5, 1.3, size=22, bold=True, color=BLUE)
txt(3, "IBM 오픈소스 양자 프레임워크", 2.5, 7.0, 13.5, 1.0, size=14, color=GRAY)
rect(3, 2.5, 8.1, 13.0, 0.04, fill=DIVIDER)
txt(3, "Python 기반 양자 회로 설계\n실제 IBM 양자 하드웨어 접근\n시뮬레이터 로컬 실행 지원\n400만+ 등록 사용자 (2024)\nApache 2.0 오픈소스 라이선스",
    2.5, 8.4, 13.5, 6.5, size=14, color=WHITE)

# 오른쪽: 응용 분야 4개
apps = [
    ("신약 개발",    "단백질 폴딩·분자 시뮬레이션\n기존 슈퍼컴퓨터 한계 극복",   BLUE),
    ("금융 최적화",  "포트폴리오 최적화·리스크 분석\n복잡한 조합 최적화 문제 해결", CYAN),
    ("암호화·보안",  "양자 내성 암호 개발\nPost-Quantum Cryptography",         "7C3AED"),
    ("물류·공급망",  "최단 경로 탐색·스케줄링\n양자 어닐링 기반 최적화",         "0EA5E9"),
]
for i, (title, desc, col) in enumerate(apps):
    row, col_off = divmod(i, 2)
    lx = 17.5 + col_off * 8.1
    ty = 5.4  + row * 5.5
    rrect(3, lx, ty, 7.5, 4.8, fill=DARK_CARD)
    rect(3,  lx, ty, 7.5, 0.3, fill=col)
    txt(3, title, lx+0.3, ty+0.5, 7.0, 1.0, size=16, bold=True, color=WHITE)
    txt(3, desc,  lx+0.3, ty+1.7, 7.0, 2.8, size=12, color=GRAY)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Slide 4 — 마무리 / IBM Quantum Network
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
print("\n[Slide 4] 마무리")
rect(4, 0, 0, 33.87, 19.05, fill="060E1A")     # 풀 다크 배경
rect(4, 0, 0, 0.55, 19.05, fill=BLUE)           # 왼쪽 강조 바

# 중앙 메인 텍스트
txt(4, "IBM Quantum Network",    2.5,  2.5, 29, 1.5, size=15, bold=True, color=BLUE,  align="center")
txt(4, "양자 기술의 미래를\n함께 만들어갑니다",  2.5,  4.0, 29, 5.5, size=44, bold=True, color=WHITE, align="center")
rect(4, 13.0, 10.0, 7.87, 0.08, fill=BLUE)

# 통계 3개
for i, (num, label) in enumerate([
    ("180+",    "글로벌 파트너사"),
    ("500,000+","Qiskit 사용자"),
    ("1T+",     "양자 연산 실행 횟수"),
]):
    lx = 3.0 + i * 9.8
    txt(4, num,   lx, 11.2, 9.0, 2.0, size=32, bold=True, color=BLUE,  align="center")
    txt(4, label, lx, 13.0, 9.0, 1.2, size=14, color=GRAY, align="center")

rect(4, 2.5, 15.8, 28.87, 0.06, fill=DIVIDER)
txt(4, "quantum.ibm.com  |  github.com/Qiskit", 2.5, 16.1, 29, 1.0, size=13, color="475569", align="center")


bridge("save_presentation")
print(f"\n=== 모든 슬라이드 완성 → {os.path.abspath(FILE)} ===")

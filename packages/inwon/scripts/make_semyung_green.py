"""
세명컴퓨터고등학교 소개 PPTX — 녹색 팔레트 · 16슬라이드 풍부한 내용
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from canva_ppt_mcp.models import (
    ChartSeries, ChartSpec, ContentItem, DeckPlan, DesignSystem,
    Palette, SlideSpec, Typography,
)
from canva_ppt_mcp.render import render_auto_deck, write_speaker_notes

OUTPUT = "output/semyung_computer_hs.pptx"

# ── 녹색 팔레트 ──────────────────────────────────────────────────────────────
green_palette = Palette(
    primary="#1B4D35",
    secondary=["#D4ECD9", "#4A8C6A"],
    accent="#F5A623",
    background_light="#F4FAF6",
    background_dark="#0E2E1E",
)
design = DesignSystem(
    palette=green_palette,
    typography=Typography(header_font="Bookman Old Style"),
    visual_motif="organic 스타일 — 녹색 자연 모티프, 비대칭 레이아웃, 대형 텍스처 배경",
    style_preset="organic",
    layout_rotation=["title", "image_focus", "icon_rows", "timeline",
                     "two_column", "chart", "grid_2x2", "comparison"],
    visual_intensity="bold",
    dark_slide_ratio=0.45,
    gradient_backgrounds=True,
    background_texture=True,
    dynamic_composition=True,
)

# ── 슬라이드 정의 (16장) ──────────────────────────────────────────────────────
slides = [

    # 1. 표지
    SlideSpec(
        title="세명컴퓨터고등학교",
        subtitle="충청북도 제천 · IT 전문 특성화 고등학교 · 1998년 개교",
        layout="title",
        speaker_notes="세명컴퓨터고등학교 학교 소개 발표입니다. IT 특성화 고등학교로서의 강점과 진로를 안내합니다.",
    ),

    # 2. 학교 연혁 & 개요
    SlideSpec(
        title="26년의 IT 교육 역사",
        subtitle="학교 개요",
        layout="timeline",
        items=[
            ContentItem(
                heading="1998년",
                body="충청북도 제천시에 세명컴퓨터고등학교 개교. 지역 IT 인재 양성 특성화 고교로 출발"),
            ContentItem(
                heading="2005년",
                body="정보통신네트워크과 신설 및 실습 인프라 대폭 확충. 산학협력 체계 구축 시작"),
            ContentItem(
                heading="2015년",
                body="스마트전자과 개설 및 메이커스페이스 구축. IoT·임베디드 실습 환경 도입"),
            ContentItem(
                heading="2024년",
                body="AI·클라우드 교육과정 전면 개편. 재학생 약 600명, 누적 졸업생 7,000명+ 배출"),
        ],
        speaker_notes="1998년 개교 이래 26년간 IT 인재를 배출해온 역사를 타임라인으로 설명합니다.",
    ),

    # 3. 학교 기본 정보
    SlideSpec(
        title="한눈에 보는 세명컴퓨터고",
        subtitle="학교 기본 현황",
        layout="grid_2x2",
        items=[
            ContentItem(
                heading="위치 · 규모",
                body="충청북도 제천시 세명로 65\n부지 2만 6천㎡ · 본관 · 실습동 · 기숙사 완비"),
            ContentItem(
                heading="재학생 현황",
                body="전교생 약 600명 (남녀공학)\n학급당 25~28명 소수정예 운영"),
            ContentItem(
                heading="교원 구성",
                body="전임교원 45명 · 산업체 겸임교사 12명\nIT 현장 경력 평균 8년 이상 전문 교원"),
            ContentItem(
                heading="운영 법인",
                body="학교법인 세명학원 산하 운영\n세명대학교와 교육·시설 연계 협약"),
        ],
        speaker_notes="학교 위치, 규모, 교원 구성 등 기본 현황을 소개합니다.",
    ),

    # 4. 학과 구성 상세
    SlideSpec(
        title="3개 전공학과 체계",
        subtitle="학과 구성",
        layout="comparison",
        items=[
            ContentItem(
                heading="컴퓨터소프트웨어과",
                body="Python·Java·C++ 프로그래밍\n앱·웹·게임 개발 프로젝트\nAI 기초·데이터 분석 실습\n정보처리기능사·SW개발 자격 취득"),
            ContentItem(
                heading="정보통신네트워크과",
                body="네트워크 설계·라우팅·스위칭\n정보보안·방화벽·침입탐지 실습\n클라우드(AWS·Azure) 실습 환경\n네트워크관리사·정보보안산업기사"),
            ContentItem(
                heading="스마트전자과",
                body="아두이노·라즈베리파이 임베디드\nIoT 센서·회로 설계 및 제작\n3D 프린팅·PCB 설계 실습\n전자CAD·전자기기기능사 취득"),
        ],
        speaker_notes="3개 학과의 교육 내용과 취득 가능 자격증을 비교 설명합니다.",
    ),

    # 5. 교육 목표
    SlideSpec(
        title="세명이 키우는 IT 인재상",
        subtitle="교육 목표",
        layout="two_column",
        items=[
            ContentItem(
                heading="실무 즉시 투입 가능한 기술력",
                body="3년간 전공 실습 1,200시간 이상 이수. 산업체 현장실습 4주 의무화로 졸업 전 현장 경험 확보"),
            ContentItem(
                heading="SW·HW 융합 문제 해결 역량",
                body="소프트웨어·하드웨어·네트워크를 통합적으로 이해하는 복합 엔지니어 양성 교육과정"),
            ContentItem(
                heading="협업과 소통 능력",
                body="팀 프로젝트 기반 수업, 발표·문서화 훈련으로 현장에서 요구하는 협업 역량을 3년간 체계적으로 키움"),
            ContentItem(
                heading="자기주도 학습 습관",
                body="방과후 심화반, 개인 포트폴리오 제작, 교내 해커톤 참가를 통해 스스로 성장하는 학습 태도 형성"),
        ],
        speaker_notes="세명컴퓨터고가 추구하는 4가지 인재상과 구체적인 교육 방식을 설명합니다.",
    ),

    # 6. 교육과정 로드맵
    SlideSpec(
        title="3년 교육 로드맵",
        subtitle="학년별 교육과정",
        layout="timeline",
        items=[
            ContentItem(
                heading="1학년 — 기초 다지기",
                body="컴퓨터 구조 이해 / 코딩 입문(Python·C) / 전기·전자 기초\n공통 교양(수학·영어·SW 리터러시) / 진로 탐색 활동"),
            ContentItem(
                heading="2학년 — 전공 심화",
                body="학과별 전공 심화 이론·실습 / 국가기술자격증 집중 취득\n산업체 멘토링·현장 견학 / 팀 프로젝트 시작\n교내 경진대회·해커톤 참가"),
            ContentItem(
                heading="3학년 — 진로 완성",
                body="캡스톤 디자인 프로젝트(개인·팀) / 산업체 현장실습 4주\n취업 면접·포트폴리오 준비 / 대학 수시 집중 지도\n국제자격증(CompTIA·AWS) 지원"),
        ],
        speaker_notes="학년별 교육과정 흐름을 타임라인으로 설명합니다.",
    ),

    # 7. 핵심 자격증 현황 — big stat
    SlideSpec(
        title="자격증 취득 실적",
        subtitle="최근 3년 평균 · 국가기술자격 기준",
        layout="big_stat",
        items=[
            ContentItem(
                heading="연간 국가자격증 취득 건수",
                body="정보처리기능사 · 네트워크관리사 · 전자CAD기능사 · 리눅스마스터 · 정보보안산업기사 · 전자기기기능사 등 20종 이상 지원",
                value="연 350+",
            ),
        ],
        speaker_notes="최근 3년 평균 기준 연간 350건 이상의 국가자격증이 재학생들에 의해 취득되고 있습니다.",
    ),

    # 8. 특색 교육 프로그램
    SlideSpec(
        title="세명만의 특색 프로그램",
        subtitle="특색 교육",
        layout="icon_rows",
        items=[
            ContentItem(
                heading="교내 해커톤 (연 2회)",
                body="48시간 팀 프로젝트 대회. 우수작은 교육부 학생 창작물 공모전 출품 및 특허 출원 지원"),
            ContentItem(
                heading="IT 창업 동아리",
                body="앱·웹 서비스 창업 아이디어를 실제 제품으로 구현. 지역 스타트업과 멘토-멘티 연결"),
            ContentItem(
                heading="세명 오픈소스 클럽",
                body="GitHub 기반 오픈소스 프로젝트 참여. 실제 커뮤니티 기여 경험으로 포트폴리오 강화"),
            ContentItem(
                heading="AI·빅데이터 특강반",
                body="방과후 심화 과정. 머신러닝·딥러닝 기초부터 kaggle 실습까지 단계적 커리큘럼 운영"),
            ContentItem(
                heading="글로벌 IT 연수",
                body="매년 2학년 희망자 대상 일본·대만 IT 기업 방문 연수. 글로벌 시각과 네트워크 형성"),
        ],
        speaker_notes="해커톤, 창업 동아리, 오픈소스 클럽 등 세명만의 차별화된 특색 프로그램을 소개합니다.",
    ),

    # 9. 산학협력 현황
    SlideSpec(
        title="산업체와 함께 성장",
        subtitle="산학협력",
        layout="grid_2x2",
        items=[
            ContentItem(
                heading="협약 기업 현황",
                body="국내외 IT 기업 80곳 이상과 산학협약 체결\nSamsung SDS · LG CNS · KT · SK텔레콤 등 대기업 포함"),
            ContentItem(
                heading="현장실습 운영",
                body="3학년 2학기 4주 현장실습 의무화\n참여 학생 95% '실무 역량 향상에 도움' 평가"),
            ContentItem(
                heading="채용 연계",
                body="협약 기업 우선 채용 트랙 운영\n졸업 전 취업 확정 비율 40% 달성 (2023년 기준)"),
            ContentItem(
                heading="겸임교사 파견",
                body="현직 IT 엔지니어 12명 겸임교사로 참여\n최신 현장 기술·트렌드를 교실에서 직접 전달"),
        ],
        speaker_notes="80여개 협약 기업과의 산학협력 체계, 현장실습, 채용 연계 현황을 소개합니다.",
    ),

    # 10. 수상 및 대외 성과
    SlideSpec(
        title="대회 수상 및 대외 성과",
        subtitle="최근 3년 주요 실적",
        layout="icon_rows",
        items=[
            ContentItem(
                heading="전국 기능경기대회",
                body="정보기술·웹디자인·IT 네트워크 직종 금·은·동 다수 수상\n충북 대표 선발 연속 6회 달성"),
            ContentItem(
                heading="SW 창작 경진대회",
                body="교육부 주관 전국 고교생 SW 경진대회 최우수상 수상 (2022·2023 연속)\n앱·AI 부문 학생 팀 수상"),
            ContentItem(
                heading="한국정보올림피아드",
                body="최근 3년 연속 입상자 배출. 입상자 전원 목표 대학 수시 합격"),
            ContentItem(
                heading="교육부 우수학교 선정",
                body="직업교육 혁신 우수학교 3년 연속 지정\n충청북도 특성화고 학력우수상 수상"),
        ],
        speaker_notes="전국 기능경기대회, SW 경진대회, 교육부 우수학교 선정 등 대외 성과를 소개합니다.",
    ),

    # 11. 졸업생 진로 차트
    SlideSpec(
        title="졸업생 진로 현황",
        subtitle="최근 3년 평균 (2021~2023 졸업생 기준)",
        layout="chart",
        items=[
            ContentItem(
                heading="진로 현황",
                body="취업·4년제·전문대·기타 비율 (최근 3년 평균)"),
        ],
        chart=ChartSpec(
            chart_type="column",
            categories=["2021년", "2022년", "2023년"],
            series=[
                ChartSeries(name="취업", values=[52, 55, 58]),
                ChartSeries(name="4년제 대학", values=[27, 25, 24]),
                ChartSeries(name="전문대학", values=[16, 15, 13]),
                ChartSeries(name="기타", values=[5, 5, 5]),
            ],
            value_suffix="%",
        ),
        speaker_notes="연도별 졸업생 진로 추이를 컬럼 차트로 보여줍니다. 취업률이 매년 증가 추세입니다.",
    ),

    # 12. 주요 취업·진학 기업 및 대학
    SlideSpec(
        title="졸업 후 향하는 곳",
        subtitle="주요 취업 기업 및 진학 대학",
        layout="two_column",
        items=[
            ContentItem(
                heading="주요 취업 기업",
                body="삼성SDS · LG CNS · SK텔레콤 · KT\n네이버 클라우드 · 카카오엔터프라이즈\n지역 IT 중소·중견 기업 다수\n공공기관 전산직 (전산직 공무원·공기업)"),
            ContentItem(
                heading="주요 진학 대학",
                body="세명대학교 (연계 특별전형)\n한국기술교육대학교 · 충북대학교\n한양대·고려대 컴퓨터공학과 (수시)\n한국폴리텍대학 (취업 후 진학)"),
            ContentItem(
                heading="특기자 전형 강점",
                body="기능경기·SW경진·올림피아드 수상 실적으로 특기자 전형 합격률 업계 최고 수준 유지"),
            ContentItem(
                heading="연계 특별전형",
                body="세명대학교와 MOU 체결 — 컴퓨터학과·정보통신공학과 정원 외 특별전형 연 10명 배정"),
        ],
        speaker_notes="졸업생이 실제로 입사한 기업과 진학한 대학교를 소개합니다.",
    ),

    # 13. 생활관 및 학교생활
    SlideSpec(
        title="안전하고 쾌적한 학교생활",
        subtitle="생활 환경",
        layout="icon_rows",
        items=[
            ContentItem(
                heading="기숙사 (생활관)",
                body="1·2·3인실 총 300실 운영. 24시간 보안관리·CCTV 설치. 주 3회 청소 서비스. 타 지역 학생 우선 배정"),
            ContentItem(
                heading="급식 및 건강",
                body="자체 조리 급식 하루 3식 제공 (기숙사생). 교내 보건실 상주 간호사 1명. 정기 건강검진 실시"),
            ContentItem(
                heading="방과후 학교",
                body="전공 심화반 · 수능 대비반 · 외국어 회화반 · 체육 동아리 등 20개 프로그램 운영"),
            ContentItem(
                heading="학생 동아리",
                body="IT 개발 동아리 8개 · 문화예술 동아리 6개 · 봉사 동아리 4개. 동아리 활동비 학교 전액 지원"),
        ],
        speaker_notes="기숙사, 급식, 방과후 학교, 동아리 등 학교생활 환경을 소개합니다.",
    ),

    # 14. 교육 시설
    SlideSpec(
        title="최첨단 실습·교육 시설",
        subtitle="시설 현황",
        layout="comparison",
        items=[
            ContentItem(
                heading="SW 개발 실습실 (4실)",
                body="고성능 PC 160대 (i9·32GB RAM)\n클라우드 개발 환경 (AWS Educate)\n3D 프린터 8대 · VR 체험 키트"),
            ContentItem(
                heading="네트워크·보안 랩 (2실)",
                body="Cisco 라우터·스위치 실장비 구비\n모의 침투 테스트 전용 망분리 환경\nSIEM·방화벽·IDS/IPS 실습 가능"),
            ContentItem(
                heading="전자·IoT 실험실 (2실)",
                body="오실로스코프·스펙트럼 분석기 보유\n아두이노·라즈베리파이 키트 80세트\nPCB 에칭 장비·납땜 스테이션"),
            ContentItem(
                heading="스마트 도서관·강당",
                body="IT 전문 도서 4,500권·전자책 무제한\n200석 대강당 (행사·특강·해커톤)\n카페형 개인학습실 50석"),
        ],
        speaker_notes="SW 실습실, 네트워크 랩, 전자 실험실, 도서관 등 시설 현황을 소개합니다.",
    ),

    # 15. 입학 전형 안내
    SlideSpec(
        title="입학 전형 안내",
        subtitle="신입생 모집 요강",
        layout="grid_2x2",
        items=[
            ContentItem(
                heading="지원 자격",
                body="중학교 졸업(예정)자 전국 누구나 지원 가능\n내신 성적 제한 없음 · 남녀 구분 없이 모집"),
            ContentItem(
                heading="전형 방법",
                body="서류 전형 (학생부 성적 40%) + 면접 (60%)\n면접 주요 항목: IT 관심도·진로 계획·협업 의지"),
            ContentItem(
                heading="모집 인원",
                body="컴퓨터소프트웨어과 60명\n정보통신네트워크과 60명\n스마트전자과 60명 / 총 180명"),
            ContentItem(
                heading="지원 일정 (2025년)",
                body="원서 접수: 12월 초 / 면접: 12월 중순\n합격 발표: 12월 말\n입학식: 2026년 3월 2일"),
        ],
        speaker_notes="입학 전형 지원 자격, 방법, 모집 인원, 일정을 안내합니다.",
    ),

    # 16. 클로징
    SlideSpec(
        title="세명에서 미래를 시작하세요",
        subtitle="IT 전문가로 성장하는 첫걸음, 세명컴퓨터고등학교가 함께합니다.",
        layout="closing",
        items=[
            ContentItem(
                heading="입학 문의 및 방문 상담",
                body="충청북도 제천시 세명로 65  |  Tel. 043-651-XXXX  |  www.semyung.hs.kr\n입학담당관 직통: 043-651-XXXX  |  카카오톡 채널: @세명컴고"),
        ],
        speaker_notes="입학 문의처와 방문 상담 안내로 마무리합니다.",
    ),
]

plan = DeckPlan(
    communication_job=(
        "발표가 끝날 때 청중(학생·학부모)은 세명컴퓨터고등학교의 학과·교육과정·산학협력·시설·진로를 구체적으로 이해하고 입학을 긍정적으로 검토한다."
    ),
    design_system=design,
    slides=slides,
    grounded=False,
    language="ko",
)

print(f"[1/2] PPTX 렌더링 중... ({len(slides)}슬라이드) -> {OUTPUT}")
Path(OUTPUT).parent.mkdir(parents=True, exist_ok=True)
render_auto_deck(plan, OUTPUT)
write_speaker_notes(OUTPUT, plan)

print(f"[2/2] 완료: {Path(OUTPUT).resolve()}")
print(f"  슬라이드 수: {len(plan.slides)}")
print(f"  스타일: {plan.design_system.style_preset}")
print(f"  팔레트 primary: {plan.design_system.palette.primary}")
for i, slide in enumerate(plan.slides, 1):
    print(f"  [{i:02d}] [{slide.layout:12s}] {slide.title}")

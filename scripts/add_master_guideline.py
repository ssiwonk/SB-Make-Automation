#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add Unified Master Guideline: "SCU 대학 강의용 보드 지침" to Notion PPT-Guidelines DB
기존 5개 지침을 조율·통합하여 대학 강의용 단일 마스터 지침으로 등록
"""

import os
import sys
import json
import urllib.request

# Windows 콘솔 UTF-8 인코딩 보장
if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

def _load_env():
    candidates = [Path.cwd() / ".env", Path(__file__).resolve().parent.parent / ".env", Path(__file__).resolve().parent / ".env"]
    for env_file in candidates:
        if env_file.exists():
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        if k.strip() not in os.environ:
                            os.environ[k.strip()] = v.strip().strip('"').strip("'")
            break

_load_env()
NOTION_TOKEN = os.environ.get("NOTION_TOKEN", "")
GUIDELINES_DB_ID = "3fa624a3-5444-80c1-9c4e-c4ea770a8a92"

NOTION_VERSION = "2022-06-28"
NOTION_API_BASE = "https://api.notion.com/v1"

def notion_request(endpoint, method="GET", body=None):
    url = f"{NOTION_API_BASE}/{endpoint.lstrip('/')}"
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Content-Type": "application/json",
        "Notion-Version": NOTION_VERSION
    }
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        print(f"[Notion API Error] {e.code}: {err_msg}", file=sys.stderr)
        raise

def p_block(text, bold_prefix=None):
    rich_text = []
    if bold_prefix:
        rich_text.append({
            "type": "text",
            "text": {"content": bold_prefix},
            "annotations": {"bold": True}
        })
    rich_text.append({
        "type": "text",
        "text": {"content": text}
    })
    return {
        "object": "block",
        "type": "paragraph",
        "paragraph": {"rich_text": rich_text}
    }

def h1_block(text):
    return {
        "object": "block",
        "type": "heading_1",
        "heading_1": {"rich_text": [{"type": "text", "text": {"content": text}}]}
    }

def h2_block(text):
    return {
        "object": "block",
        "type": "heading_2",
        "heading_2": {"rich_text": [{"type": "text", "text": {"content": text}}]}
    }

def h3_block(text):
    return {
        "object": "block",
        "type": "heading_3",
        "heading_3": {"rich_text": [{"type": "text", "text": {"content": text}}]}
    }

def bullet_block(text, bold_prefix=None):
    rich_text = []
    if bold_prefix:
        rich_text.append({
            "type": "text",
            "text": {"content": bold_prefix},
            "annotations": {"bold": True}
        })
    rich_text.append({
        "type": "text",
        "text": {"content": text}
    })
    return {
        "object": "block",
        "type": "bulleted_list_item",
        "bulleted_list_item": {"rich_text": rich_text}
    }

def callout_block(text, emoji="🎓"):
    return {
        "object": "block",
        "type": "callout",
        "callout": {
            "icon": {"type": "emoji", "emoji": emoji},
            "rich_text": [{"type": "text", "text": {"content": text}}]
        }
    }

def code_block(code_str, language="markdown"):
    return {
        "object": "block",
        "type": "code",
        "code": {
            "language": language,
            "rich_text": [{"type": "text", "text": {"content": code_str}}]
        }
    }

def divider_block():
    return {"object": "block", "type": "divider", "divider": {}}

# ==========================================
# SCU 대학 강의용 통합 마스터 지침 블록 정의
# ==========================================
MASTER_TITLE = "SCU 대학 강의용 보드 지침"
MASTER_TYPES = ["통합 마스터", "SCU 표준", "대학강의"]

MASTER_BLOCKS = [
    callout_block(
        "본 지침은 서울사이버대학교(SCU)의 정규 대학 강의용 슬라이드 및 스토리보드(SB) 제작에 필요한 "
        "모든 표준(원고 정제, 규격, 타이포그래피, 인포그래픽 시각화, SVG 엔지니어링, 품질 검수)을 "
        "단일 체계로 통합한 공식 마스터 지침입니다.",
        "🎓"
    ),
    divider_block(),

    # 1. 원고 정제 및 마크다운 표준
    h2_block("1. 원고 정제 및 슬라이드 구조화 표준"),
    p_block("Marp/Markdown 기반 원고 작성 및 슬라이드 분할 규칙:"),
    bullet_block(" 최상단에 표지 단 1개만 사용 (과목명, 주차, 차시, 담당교수명).", "h1 (표지):"),
    bullet_block(" 대단원/유닛 전환 지점에만 1회 표기 (동일 유닛 내 반복 금지).", "h2 (간지):"),
    bullet_block(" 개별 슬라이드의 대표 제목. 독립 섹션(형성평가, 학습정리 등)도 h3 적용.", "h3 (슬라이드 제목):"),
    bullet_block(" 슬라이드 핵심 1줄 요약. 원문에 없을 경우 본문을 근거로 AI가 자동 도출.", "h4 (가버링 소제목):"),
    bullet_block("h4는 단독 존재 불가(반드시 상위 h3 동반). 각 슬라이드 블록 끝에는 빈 줄과 '---'를 삽입하여 슬라이드 경계 확정.", "동반 및 분할 규칙:"),
    bullet_block("텍스트 본문은 유지하되 볼드/기울임 마크업 및 '유닛N', '페이지N', '슬라이드 N.', 시간코드, '[요약]', '※ 화면 구성:' 등 개발/연출 지시어 100% 삭제.", "지시어 제거:"),
    divider_block(),

    # 2. 캔버스 규격 및 타이포그래피 조율 기준
    h2_block("2. 캔버스 규격 및 타이포그래피 조율 기준"),
    callout_block(
        "규격 조율: 기본 4:3(1024×768)으로 제작 후 16:9 와이드(13.333×7.5 inch) 변환본(-169.pptx)을 자동 생성하여 두 규격을 모두 지원합니다.",
        "📐"
    ),
    bullet_block("4:3 표준 (1024×768 px), 좌우 60px 안전 마진 (가용 폭 904px, x=60~964).", "화면 규격:"),
    bullet_block("42pt 이상 ➔ SVG 56px ~ 64px (Bold, 텍스트 기본색)", "h1 (표지):"),
    bullet_block("32pt 이상 ➔ SVG 42px ~ 48px (Bold, 중앙 정렬)", "h2 (간지):"),
    bullet_block("32pt ➔ SVG 42.67px ~ 43px (Bold, 텍스트 기본색)", "h3 (슬라이드 제목):"),
    bullet_block("24pt ➔ SVG 32px (Semi-bold, 주색상/포인트색, 긴 제목 2줄 분할 시 최소 28px)", "h4 (가버링 소제목):"),
    bullet_block("20pt 이상 ➔ SVG 약 26.67px 권장 (SCU 정량 기준 충족, 최소 16pt/21.33px 이상)", "본문 카드/설명:"),
    bullet_block("최소 14pt 이상 ➔ SVG 18.67px (인포그래픽 내부 세부 라벨 및 캡션 한정)", "다이어그램 세부라벨:"),
    bullet_block("헤더 텍스트(h1~h4)는 토씨 하나 바꾸지 않고 100% 보존. 임의 키커, 페이지 번호, 임의 푸터 배제.", "텍스트 무결성:"),
    divider_block(),

    # 3. 70% 인포그래픽 도해화 및 슬라이드 분할 규칙
    h2_block("3. 70% 인포그래픽 시각화 및 슬라이드 분할(Split) 규칙"),
    bullet_block("단순 줄글 나열을 엄격히 지양하고 슬라이드 면적의 70% 이상을 다이어그램 및 구조화된 카드로 구성.", "인포그래픽 극대화:"),
    bullet_block("원형 플라이휠(선순환), 프로세스 셰브론(단계별 로드맵), 피라미드 스택(위계/계층), 대칭형 밸런스 카드(비교/대조), 매트릭스 그리드(병렬 시책), 분기 트리(전달체계).", "6대 핵심 도해 패턴:"),
    bullet_block("한글 본문 폰트 기준 가로 800px 박스당 34자 내외로 래핑하여 오버플로우 원천 차단.", "도형 이탈 방지:"),
    bullet_block("법률 조문, 복수 사례, 긴 원문 등 정보과밀 우려 시, 동일한 h3, h4 제목을 유지한 채 슬라이드를 2장 이상(①, ②)으로 분할(Split)하여 가독성 확보.", "슬라이드 분할(Split):"),
    divider_block(),

    # 4. SVG 엔지니어링 및 품질 검사(QA) 규격
    h2_block("4. SVG 엔지니어링 및 렌더링 품질 규격"),
    bullet_block("<svg> 바로 아래 모든 <g> 태그는 반드시 'data-pptx-bounds=\"x y w h\"' 명시.", "Root 바운즈 필수:"),
    bullet_block("배경 객체는 반드시 id=\"slide-bg\" 및 data-pptx-role=\"background\" 속성 부여.", "배경 ID 고유화:"),
    bullet_block("본문 텍스트 내 &, <, > 등은 &amp;, &lt;, &gt; 로 완벽히 이스케이프.", "XML 특수문자:"),
    bullet_block("h4가 28자 이상일 경우 <tspan x=\"60\" dy=\"34\">로 2줄 분할 배치.", "긴 텍스트 tspan:"),
    bullet_block("원고 내 이미지(![...](경로))는 슬라이드 레이아웃(좌측 이미지 + 우측 인포그래픽 등)에 1:1 필수 반영.", "이미지 1:1 배치:"),
    divider_block(),

    # 5. SCU 2-Track 품질 검수 및 실무 지시어 작성 원칙
    h2_block("5. SCU 2-Track 품질 검수 및 쉬운 실무 지침투 작성 원칙"),
    bullet_block("슬라이드 규격, 본문 폰트 20pt+ Bold, 화면 경계 오버플로우, 슬라이드 노트 내 '#출처' 누락, 마침표 오표기 검사.", "Track 1 (정량 룰 검사):"),
    bullet_block("원고 vs SB 왜곡/누락 대조, 정보과밀 감지(분할 제안), 학습목표 도달 행동 진술, 도해화 제안.", "Track 2 (정성 AI 검수):"),
    bullet_block("추상적인 교수설계 용어 절대 금지. 파워포인트에서 1초 만에 바로 손을 움직일 수 있는 직접 지시어 작성.", "쉬운 실무 지시어 원칙:"),
    bullet_block(" ❌ '핵심 행동목표 모호' ➔ ⭕ \"'~을 이해할 수 있다'를 '~을 설명할 수 있다'로 바꿔주세요.\"", "학습목표 지적:"),
    bullet_block(" ❌ '대비 구조 도해화 제안' ➔ ⭕ \"아래에 박스 2개를 나란히 만들고 [왼쪽: 판매장소 X], [오른쪽: 판매전략 O]로 명확히 대비되게 넣어주세요.\"", "도해화 지적:"),
    bullet_block(" ❌ '학습정리 내용 부재' ➔ ⭕ \"학습정리가 비어있습니다. '1. 글로벌 플랫폼 6대 유형 이해'와 같이 바로 복사해 넣을 수 있는 문장을 직접 적어주세요.\"", "학습정리 지적:"),
    bullet_block("사용자가 검수 기준 수정을 요청할 때는 반드시 \"① 모든 과목 공통 적용 전사 기준인가요?\" vs \"② 현재 과목에만 적용할 고유 특화 기준인가요?\"를 먼저 확인하고 분기 반영.", "기준 보완 시 범위 확인:")
]

def add_master_guideline():
    print("=" * 70)
    print(f" [Notion PPT-Guidelines] '{MASTER_TITLE}' 등록 시작")
    print("=" * 70)

    properties = {
        "개발지침": {
            "title": [{"type": "text", "text": {"content": MASTER_TITLE}}]
        },
        "유형": {
            "multi_select": [{"name": t} for t in MASTER_TYPES]
        }
    }

    create_body = {
        "parent": {"database_id": GUIDELINES_DB_ID},
        "properties": properties,
        "children": MASTER_BLOCKS
    }

    res = notion_request("pages", method="POST", body=create_body)
    print(f" ✅ '{MASTER_TITLE}' 페이지 생성 완료! (ID: {res['id']})")
    print("=" * 70)

if __name__ == "__main__":
    add_master_guideline()

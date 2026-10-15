#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Seed Guidelines to Notion PPT-Guidelines Database
기존 워크스페이스 내 핵심 지침(spec/md-definition.md, spec/pptx-definition.md, SB-Design, AGENTS.md)을
모듈화하여 Notion PPT-Guidelines DB에 체계적으로 등록하는 스크립트
"""

import os
import sys
import json
import urllib.request
from datetime import datetime

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

def callout_block(text, emoji="💡"):
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
# 5대 모듈 지침 정의
# ==========================================
GUIDELINES_DATA = [
    {
        "title": "[원고-구조화] 슬라이드 헤더 위계 및 텍스트 정제 지침",
        "types": ["원고 정제", "구조화"],
        "blocks": [
            callout_block("Marp/Markdown 기반 원고를 슬라이드 단위로 정제할 때 반드시 준수해야 하는 헤더 위계 및 불필요한 지시어 제거 규칙입니다.", "📌"),
            h2_block("1. Marp 표준 Frontmatter"),
            p_block("문서 최상단에는 항상 아래 Marp 헤더를 포함합니다:"),
            code_block("---\nmarp: true\ntheme: my-theme\n---", "markdown"),
            divider_block(),
            h2_block("2. 4단계 헤더 위계 (h1~h4 이외 금지)"),
            bullet_block(" 문서에 단 1개만 존재. 최상단 표지 콘텐츠(과목명, 주차, 차시, 발표자)를 지정.", "h1 (표지 대제목):"),
            bullet_block(" 대단원/유닛 전환 지점에만 1회 사용. 동일 유닛 내 반복 금지.", "h2 (간지/섹션):"),
            bullet_block(" 개별 슬라이드의 대표 제목. 독립 섹션(형성평가, 학습정리 등)도 h3 적용.", "h3 (슬라이드 대제목):"),
            bullet_block(" 슬라이드의 핵심 한 줄 가버링 메시지. 원문에 없을 경우 본문을 근거로 AI가 1줄 요약 생성.", "h4 (가버링 소제목):"),
            divider_block(),
            h2_block("3. h3-h4 동반 규칙 및 슬라이드 구분자"),
            bullet_block("h4는 단독으로 존재할 수 없으며 반드시 상위에 h3가 있어야 합니다. h3 하위에 복수의 h4가 있으면 h3 제목을 동일하게 반복 표기합니다."),
            bullet_block("각 슬라이드 블록(h3 + h4 + 본문)의 끝에는 빈 줄과 함께 '---' 구분자를 삽입하여 슬라이드를 분할합니다."),
            divider_block(),
            h2_block("4. 서식 및 불필요한 개발 지시어 100% 제거"),
            bullet_block("텍스트 본문은 유지하되 볼드체(**...**) 및 기울임체(*...*) 마크업은 완전히 제거합니다."),
            bullet_block("'유닛N', '페이지N', '슬라이드 N.', 시간코드([00:00~01:00]), '[요약]', '※ 화면 구성:' 등 개발/연출용 라벨과 지시어는 모두 삭제합니다.")
        ]
    },
    {
        "title": "[디자인-규격] 캔버스 규격 및 타이포그래피(폰트 환산) 표준",
        "types": ["규격", "타이포그래피"],
        "blocks": [
            callout_block("4:3 기본 규격(1024×768) 및 16:9 와이드 변환 기준, 96 DPI 환산 타이포그래피 공식 표준 규격입니다.", "📐"),
            h2_block("1. 캔버스 규격 및 와이드 변환"),
            bullet_block("4:3 표준 비율 (1024×768 px, 10×7.5 inch), viewBox=\"0 0 1024 768\", font-family=\"Malgun Gothic\"", "기본 비율:"),
            bullet_block("좌우 60px 마진 유지 (x=60 ~ x=964, 실제 가용 너비 904px)", "안전 가용 영역:"),
            bullet_block("4:3 PPTX 제작 직후 scripts/convert-ratio.py를 실행하여 16:9 와이드 규격(13.333×7.5 inch) 자동 산출", "16:9 자동 확장:"),
            divider_block(),
            h2_block("2. 타이포그래피 환산표 (1pt = 1.3333px, 96 DPI)"),
            bullet_block(" 42pt 이상 ➔ SVG 56px ~ 64px (Bold, 지정 텍스트 기본색)", "h1 (표지):"),
            bullet_block(" 32pt 이상 ➔ SVG 42px ~ 48px (Bold, 중앙 정렬)", "h2 (간지):"),
            bullet_block(" 32pt ➔ SVG 42.67px ~ 43px (Bold, 지정 텍스트 기본색)", "h3 (슬라이드 제목):"),
            bullet_block(" 24pt ➔ SVG 32px (Semi-bold, 주색상/포인트색, 긴 제목 2줄 분할 시 최소 28px)", "h4 (가버링 소제목):"),
            bullet_block(" 16pt 이상 ➔ SVG 최소 21.33px ~ 22px (가독성 본문색)", "본문 카드/설명:"),
            bullet_block(" 14pt 이상 ➔ SVG 최소 18.67px ~ 19px (짧은 라벨 한정)", "다이어그램 캡션:"),
            divider_block(),
            h2_block("3. 텍스트 무결성 원칙"),
            bullet_block("원고의 헤더 텍스트(h1~h4)는 토씨 하나 바꾸지 않고 100% 보존하며, 넘칠 경우 줄바꿈(<tspan>)을 적용합니다."),
            bullet_block("원고에 없는 키커(kicker), 임의 페이지 번호, 임의 푸터는 일체 삽입하지 않습니다.")
        ]
    },
    {
        "title": "[레이아웃-도해화] 70% 인포그래픽 시각화 및 슬라이드 분할 규칙",
        "types": ["레이아웃", "인포그래픽"],
        "blocks": [
            callout_block("줄글 나열을 엄격히 지양하고 슬라이드 면적의 70% 이상을 다이어그램과 인포그래픽으로 구성하는 핵심 시각화 지침입니다.", "📊"),
            h2_block("1. 6대 핵심 인포그래픽 패턴"),
            bullet_block(" 핵심 이념, 유기적 선순환 생태계, 상호작용 표현", "1) 원형 플라이휠 / 도넛 링:"),
            bullet_block(" 연간 일정, 단계별 추진 절차, 시계열 연혁, 파이프라인", "2) 프로세스 체인 (수평/수직 셰브론):"),
            bullet_block(" 법체계 위계, 상·하위 법령 구조, 기술 및 인프라 스택 계층", "3) 다층 플랫폼 스택 / 피라미드:"),
            bullet_block(" 3대 핵심 주체, 권리-책임 대칭, A vs B 비교/대조 구조", "4) 대칭형 밸런스 / 3축 카드:"),
            bullet_block(" 다분야 병렬 시책, 영역별 핵심 지표, 3단 카드 레이아웃", "5) 매트릭스 타일 그리드 (2×2, 3-Column):"),
            bullet_block(" 중앙조정위 ↔ 지방조정위 이원화, 위임·위탁 전달체계", "6) 분기 네트워크 트리:"),
            divider_block(),
            h2_block("2. 도형 이탈 방지 및 슬라이드 분할 (Split) 규칙"),
            bullet_block("한글 본문 폰트(22px) 기준 너비 800px 박스에는 한 줄에 34자 이내로 배치하거나 복수 행으로 줄바꿈합니다."),
            bullet_block("법률 조문, 복수 사례, 용어 정의 등 내용이 많아 캔버스 이탈이나 가독성 저하가 우려되는 경우, 동일한 h3와 h4 제목을 그대로 유지한 채 슬라이드를 2장 이상(①, ②)으로 분할(Split)하여 제작합니다.")
        ]
    },
    {
        "title": "[SVG-엔지니어링] ppt-master SVG 계층 구조 및 품질 검사(QA) 규격",
        "types": ["엔지니어링", "품질"],
        "blocks": [
            callout_block("ppt-master 품질 검사기(svg_quality_checker.py)를 무결점(0 Errors)으로 통과하기 위한 SVG 기술 표준입니다.", "⚙️"),
            h2_block("1. 표준 슬라이드 SVG 계층 템플릿"),
            code_block(
'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 768" font-family="Malgun Gothic" lang="ko-KR" data-pptx-page-role="content">\n'
'  <!-- 1. 배경 (id="slide-bg", data-pptx-role="background" 필수) -->\n'
'  <rect id="slide-bg" width="1024" height="768" fill="{지정_배경색}" data-pptx-role="background"/>\n\n'
'  <!-- 2. 헤더 그룹 (data-pptx-bounds 필수) -->\n'
'  <g id="header-group" data-pptx-bounds="60 30 904 125">\n'
'    <text x="60" y="70" font-size="43" font-weight="700" fill="{텍스트색}">{h3 제목}</text>\n'
'    <text x="60" y="115" font-size="32" font-weight="600" fill="{포인트색}">{h4 소제목}</text>\n'
'    <rect x="60" y="142" width="904" height="2" fill="{구분선색}"/>\n'
'  </g>\n\n'
'  <!-- 3. 메인 콘텐츠 그룹 (모든 카드/도형 배치) -->\n'
'  <g id="content-group" data-pptx-bounds="60 160 904 550">\n'
'    <!-- 인포그래픽 카드, 다이어그램, 텍스트 요소들 -->\n'
'  </g>\n'
'</svg>',
                "xml"
            ),
            divider_block(),
            h2_block("2. 필수 준수 엔지니어링 규칙"),
            bullet_block("<svg> 바로 아래의 모든 <g> 태그는 반드시 'data-pptx-bounds=\"x y w h\"' 속성을 명시해야 합니다."),
            bullet_block("배경 객체는 반드시 id=\"slide-bg\" 및 data-pptx-role=\"background\" 속성을 갖추어야 합니다."),
            bullet_block("텍스트 내부의 특수문자는 XML 안전 문자(&amp;, &lt;, &gt;)로 완벽히 이스케이프합니다."),
            bullet_block("h4가 28자 이상일 경우 <tspan x=\"60\" dy=\"34\">로 2줄 분할하여 1024px 범위를 초과하지 않도록 합니다."),
            bullet_block("원고에 포함된 이미지(![...](경로))는 슬라이드 레이아웃(좌측 이미지 + 우측 인포그래픽 등)에 1:1로 정확히 반영합니다.")
        ]
    },
    {
        "title": "[검수-품질] SCU 교육용 SB 품질 기준 및 쉬운 실무 지침투 작성 원칙",
        "types": ["검수기준", "톤앤매너"],
        "blocks": [
            callout_block("현업 제작자가 보고서를 보자마자 고민 없이 즉시 파워포인트에서 손을 움직일 수 있는 실무 지시어 작성 원칙입니다.", "🎯"),
            h2_block("1. 2-Track 품질 검수 체계"),
            bullet_block(" 슬라이드 규격(16:9/4:3), 본문 폰트 20pt+ Bold, 화면 경계 오버플로우, 슬라이드 노트 내 '#출처' 누락, 마침표 오표기 검사", "Track 1 (정량 룰 검사):"),
            bullet_block(" 원고 vs SB 내용 누락/왜곡 대조, 정보과밀 슬라이드 2장 분할 제안, 학습목표 도달 행동 진술, 대비형 도해화 제안", "Track 2 (정성 AI 검수):"),
            divider_block(),
            h2_block("2. 영구기억 규칙: 쉬운 실무 지시어 (Before vs After)"),
            bullet_block(" ❌ '핵심 행동목표가 모호하므로 도달 행동을 명확히 진술 권고'\n ⭕ \"'~을 이해할 수 있다'를 '~을 설명할 수 있다'로 바꿔주세요.\"", "학습목표 지적 시:"),
            bullet_block(" ❌ '텍스트 나열을 지양하고 판매장소 vs 판매전략 대비 구조로 도해화 제안'\n ⭕ \"텍스트 한 줄만 덜렁 두지 마세요. 아래에 박스 2개를 나란히 만들고 [왼쪽: 판매장소 선택 X], [오른쪽: 판매전략 선택 O]로 명확히 대비되게 넣어주세요.\"", "도해화 지적 시:"),
            bullet_block(" ❌ '서브타이틀 일관성 및 대등 구조 유지 검토'\n ⭕ \"슬라이드 위쪽 서브타이틀 박스에 'Amazon에서 중요한 질문'이라고 제목을 붙여주세요.\"", "제목/서브타이틀 지적 시:"),
            bullet_block(" ❌ '학습정리 내용 부재로 지침서 미준수'\n ⭕ \"학습정리가 비어있습니다. 비워두지 마시고 '1. 글로벌 플랫폼 6대 유형 이해, 2. 주요 플랫폼 비교'와 같이 바로 복사해 넣을 수 있는 문장을 직접 적어주세요.\"", "학습정리 지적 시:"),
            divider_block(),
            h2_block("3. 검수 기준 보완 시 반영 범위 확인 원칙"),
            p_block("사용자가 검수 기준 수정을 요청할 때는 반드시 아래 양식으로 먼저 질문한 뒤 분기 반영합니다:"),
            callout_block("\"이 기준은 ① 모든 과목에 공통으로 적용할 전사 기준인가요, 아니면 ② 현재 과목({과목명})에만 적용할 고유 특화 기준인가요?\"", "❓")
        ]
    }
]

def seed_guidelines():
    print("=" * 70)
    print(" [Notion PPT-Guidelines] 기존 지침 체계적 등록 시작")
    print("=" * 70)
    
    # 기존에 생성되어 있던 빈 페이지가 있는지 확인
    db_items = notion_request(f"databases/{GUIDELINES_DB_ID}/query", method="POST", body={"page_size": 100})
    existing_pages = db_items.get("results", [])
    
    for idx, g_data in enumerate(GUIDELINES_DATA):
        title = g_data["title"]
        types = g_data["types"]
        blocks = g_data["blocks"]
        
        properties = {
            "개발지침": {
                "title": [{"type": "text", "text": {"content": title}}]
            },
            "유형": {
                "multi_select": [{"name": t} for t in types]
            }
        }
        
        # 첫 번째 항목이고 기존에 빈 페이지가 있으면 업데이트로 재활용
        if idx == 0 and existing_pages:
            reuse_page = existing_pages[0]
            p_id = reuse_page["id"]
            print(f"* 기존 빈 페이지({p_id})를 [{title}]로 업데이트...")
            notion_request(f"pages/{p_id}", method="PATCH", body={"properties": properties})
            # 블록 추가
            notion_request(f"blocks/{p_id}/children", method="PATCH", body={"children": blocks})
            print(f"  -> 등록 완료: {title}")
        else:
            print(f"* 신규 지침 등록: [{title}]...")
            create_body = {
                "parent": {"database_id": GUIDELINES_DB_ID},
                "properties": properties,
                "children": blocks
            }
            res = notion_request("pages", method="POST", body=create_body)
            print(f"  -> 등록 완료: {title} (ID: {res['id']})")
            
    print("\n" + "=" * 70)
    print(" ✅ 총 5개 핵심 지침 모듈이 PPT-Guidelines에 성공적으로 등록되었습니다!")
    print("=" * 70)

if __name__ == "__main__":
    seed_guidelines()

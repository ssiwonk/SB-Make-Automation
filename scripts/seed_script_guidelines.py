#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Seed Script Transformation Guidelines to Notion PPT-Guidelines DB
원고 변환(HWP/DOCX/PDF/PPTX -> Markdown)을 위한 4대 핵심 지침 모듈 등록
"""

import os
import sys
import json
import urllib.request

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
        rich_text.append({"type": "text", "text": {"content": bold_prefix}, "annotations": {"bold": True}})
    rich_text.append({"type": "text", "text": {"content": text}})
    return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": rich_text}}

def h2_block(text):
    return {"object": "block", "type": "heading_2", "heading_2": {"rich_text": [{"type": "text", "text": {"content": text}}]}}

def bullet_block(text, bold_prefix=None):
    rich_text = []
    if bold_prefix:
        rich_text.append({"type": "text", "text": {"content": bold_prefix}, "annotations": {"bold": True}})
    rich_text.append({"type": "text", "text": {"content": text}})
    return {"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": rich_text}}

def callout_block(text, emoji="📝"):
    return {"object": "block", "type": "callout", "callout": {"icon": {"type": "emoji", "emoji": emoji}, "rich_text": [{"type": "text", "text": {"content": text}}]}}

def code_block(code_str, language="markdown"):
    return {"object": "block", "type": "code", "code": {"language": language, "rich_text": [{"type": "text", "text": {"content": code_str}}]}}

def divider_block():
    return {"object": "block", "type": "divider", "divider": {}}

SCRIPT_GUIDELINES = [
    {
        "title": "[원고지침] 대본형 원고 ➔ 슬라이드 개조식 추출 및 대본(Notes) 분리 규칙",
        "types": ["원고변환", "대본형", "슬라이드노트"],
        "blocks": [
            callout_block("강의 녹취록이나 구어체 대본 형태의 원고를 슬라이드용 화면 텍스트와 발표 대본(슬라이드 노트)으로 분리 정제하는 지침입니다.", "🎙️"),
            h2_block("1. 대본 정제 및 슬라이드 본문 추출"),
            bullet_block(" '여러분 안녕하십니까', '다음으로 넘어가겠습니다' 등 구어체 접속사 및 인사말을 100% 필터링합니다.", "구어체 제거:"),
            bullet_block(" 한 슬라이드당 3~4개의 핵심 명제(개조식 문장)로 압축하여 슬라이드 화면 텍스트로 배치합니다.", "화면 텍스트:"),
            bullet_block(" 교수자가 발화한 상세 원문 대본은 '<!-- Note: ... -->' 형태로 슬라이드 하단에 온전히 보존하여 슬라이드 노트(Notes)로 이관합니다.", "대본 분리(Notes):"),
            divider_block(),
            h2_block("2. 마크다운 출력 포맷 예시"),
            code_block(
"### 클라우드 컴퓨팅의 개념\n"
"#### 인터넷을 통해 IT 리소스를 온디맨드로 제공하는 서비스 모델\n\n"
"* 물리적 서버 구축 없이 필요한 만큼 인프라 즉시 대여\n"
"* 초기 투자 비용 절감 및 사용량 기반 종량제 과금 체계\n"
"* 글로벌 리전 기반의 신속한 서비스 배포 및 탄력적 확장성 확보\n\n"
"<!-- Note:\n"
"여러분, 클라우드 컴퓨팅이란 단어 많이 들어보셨죠? 과거에는 서버를 직접 사서 전산실에 넣었지만,\n"
"지금은 인터넷만 연결되면 아마존이나 구글의 데이터센터를 빌려 쓰는 시대가 되었습니다.\n"
"-->\n\n"
"---",
                "markdown"
            )
        ]
    },
    {
        "title": "[원고지침] 교재/논문 줄글 ➔ 핵심 요약 및 슬라이드 분할 규칙",
        "types": ["원고변환", "교재형", "줄글요약"],
        "blocks": [
            callout_block("한글(HWP)/워드(DOCX) 파일에 빼곡히 작성된 줄글 문단 원고를 슬라이드 단위(One Idea per Slide)로 분할하고 개조식으로 요약하는 지침입니다.", "📖"),
            h2_block("1. 슬라이드 분할 및 요약 원칙"),
            bullet_block(" 1개 슬라이드에는 오직 1개의 핵심 주제/개념만 배치합니다 (One Idea per Slide).", "슬라이드 분할:"),
            bullet_block(" 긴 줄글 문단을 그대로 복사하지 않고, 주어+술어가 명확한 3~4줄의 핵심 개조식 문장으로 요약 재구성합니다.", "개조식 요약:"),
            bullet_block(" 본문의 핵심 내용을 관통하는 1줄 핵심 소제목(h4 가버링 메시지)을 반드시 AI가 도출하여 부여합니다.", "가버링 메시지(h4):"),
            bullet_block(" 법률 조문, 규정, 복수 사례 등 정보량이 방대한 경우 원문을 훼손하지 않고 슬라이드를 2장 이상(①, ②)으로 분할합니다.", "정보과밀 방지:")
        ]
    },
    {
        "title": "[원고지침] 기존 PPTX/개조식 ➔ 원문 100% 직역 매핑 규칙",
        "types": ["원고변환", "직역형", "개조식"],
        "blocks": [
            callout_block("교수자가 이미 슬라이드 단위로 불릿 포인트를 정리해 둔 PPTX 또는 개조식 원고를 무손실로 마크다운에 이식하는 지침입니다.", "🎯"),
            h2_block("1. 원문 100% 보존 원칙 (Literal Mapping)"),
            bullet_block(" 원고에 작성된 단어, 문장, 수치 데이터를 임의로 축약하거나 삭제하지 않고 100% 원문 그대로 반영합니다.", "텍스트 무결성:"),
            bullet_block(" 기존 슬라이드의 페이지 제목을 h3, 소제목을 h4, 하위 불릿을 본문 리스트로 정확히 1:1 매핑합니다.", "헤더 매핑:"),
            bullet_block(" '슬라이드 1', '유닛 2', 페이지 번호, 시간코드 등 개발용 꼬리표 라벨만 선별적으로 제거합니다.", "지시어 정제:"),
            bullet_block(" 원고에 포함된 표(Table)와 이미지 경로는 표준 마크다운 표 및 이미지 링크로 무결하게 보존합니다.", "표 및 이미지 보존:")
        ]
    },
    {
        "title": "[원고지침] SCU 교수설계 표준 (도입-학습목표-본문-퀴즈-정리) 구조화 규칙",
        "types": ["원고변환", "SCU 표준", "교수설계"],
        "blocks": [
            callout_block("서울사이버대학교(SCU)의 정규 강의 표준 교수설계 5단계 흐름에 맞추어 원고의 순서와 필수 요소를 강제 구조화하는 지침입니다.", "🏛️"),
            h2_block("1. SCU 5단계 표준 슬라이드 시퀀스"),
            bullet_block(" 1) [도입] 생각해보기 또는 전시학습 연계", "1단계:"),
            bullet_block(" 2) [학습목표] '~을 설명할 수 있다', '~을 분석할 수 있다' 등 명확한 도달 행동 진술 (단순 '이해한다' 금지)", "2단계:"),
            bullet_block(" 3) [본문 학습] 소주제별 3~5개 유닛 본문 슬라이드 (70% 도해화 적용)", "3단계:"),
            bullet_block(" 4) [형성평가] 객관식/OX 퀴즈 1~2문항 + 정답 및 해설 블록 필수 배치", "4단계:"),
            bullet_block(" 5) [학습정리] 이번 차시의 핵심 학습 내용 3대 포인트 요약", "5단계:"),
            divider_block(),
            h2_block("2. 누락 요소 자동 보정"),
            p_block("교수자 원고에 학습목표나 형성평가가 누락되어 있을 경우, AI가 본문 내용을 바탕으로 표준 양식의 슬라이드 블록을 제안 형태로 자동 보강합니다.")
        ]
    }
]

def seed():
    print("=" * 70)
    print(" [Notion PPT-Guidelines] 원고 변환 지침 4종 등록 시작")
    print("=" * 70)
    for g in SCRIPT_GUIDELINES:
        title = g["title"]
        types = g["types"]
        blocks = g["blocks"]
        
        properties = {
            "개발지침": {"title": [{"type": "text", "text": {"content": title}}]},
            "유형": {"multi_select": [{"name": t} for t in types]}
        }
        create_body = {
            "parent": {"database_id": GUIDELINES_DB_ID},
            "properties": properties,
            "children": blocks
        }
        res = notion_request("pages", method="POST", body=create_body)
        print(f" ✅ 등록 완료: {title} (ID: {res['id']})")
    print("=" * 70)

if __name__ == "__main__":
    seed()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Seed 5 University Lecture Visual Styles to Notion PPT-Guidelines DB
대학 강의 5대 핵심 비주얼 스타일 지침 등록 스크립트
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

def callout_block(text, emoji="🎨"):
    return {"object": "block", "type": "callout", "callout": {"icon": {"type": "emoji", "emoji": emoji}, "rich_text": [{"type": "text", "text": {"content": text}}]}}

def code_block(code_str, language="css"):
    return {"object": "block", "type": "code", "code": {"language": language, "rich_text": [{"type": "text", "text": {"content": code_str}}]}}

def divider_block():
    return {"object": "block", "type": "divider", "divider": {}}

STYLES_DATA = [
    {
        "title": "[스타일-모던교육] 소프트 라운디드 (Soft-Rounded)",
        "types": ["디자인스타일", "테마", "모던교육"],
        "blocks": [
            callout_block("가장 호불호 없이 눈이 편안한 표준 대학 강의용 디자인 스타일입니다. 둥근 모서리와 부드러운 여백으로 친근하고 전문적인 인상을 줍니다.", "🧸"),
            h2_block("1. 디자인 특성 및 색상 가이드"),
            bullet_block(" 모서리 반경(rx/ry) 16~24px의 부드러운 라운디드 카드 서피스 적용.", "쉐입(Shape):"),
            bullet_block(" 웜 라이트 그레이 / 크림 화이트 (#F8FAFC, #FFFFFF)", "캔버스 배경:"),
            bullet_block(" 신뢰감을 주는 딥 네이비 (#0F172A)", "텍스트 기본색:"),
            bullet_block(" 소프트 인디고 블루 (#4F46E5) 및 라벤더 (#6366F1)", "주색상(Primary):"),
            bullet_block(" 코랄 오렌지 (#F97316) 또는 에메랄드 (#10B981)", "포인트색(Accent):"),
            divider_block(),
            h2_block("2. 추천 대상 과목"),
            bullet_block(" 인문학, 사회과학, 교육학, 상담심리, 교양 및 대부분의 일반 학부 강의")
        ]
    },
    {
        "title": "[스타일-스마트IT] 다크 테크 (Dark-Tech)",
        "types": ["디자인스타일", "테마", "스마트IT"],
        "blocks": [
            callout_block("딥 블랙/네이비 캔버스 위에 네온 시안 하이라이트가 돋보이는 미래지향적 테크 스타일입니다. 코드와 아키텍처 다이어그램 몰입도가 뛰어납니다.", "⚡"),
            h2_block("1. 디자인 특성 및 색상 가이드"),
            bullet_block(" 날렵하고 정밀한 직각 또는 미세 라운드(rx=6px), 1px 서피스 테두리선.", "쉐입(Shape):"),
            bullet_block(" 딥 옵시디언 블랙 / 다크 네이비 (#0B0F19, #111827)", "캔버스 배경:"),
            bullet_block(" 높은 시인성의 쿨 화이트 (#F3F4F6)", "텍스트 기본색:"),
            bullet_block(" 일렉트릭 사이언 (#00E5FF) 및 네온 블루 (#3B82F6)", "주색상(Primary):"),
            bullet_block(" 비비드 라임 (#10B981) 또는 마젠타 (#EC4899)", "포인트색(Accent):"),
            divider_block(),
            h2_block("2. 추천 대상 과목"),
            bullet_block(" 인공지능(AI), 파이썬 프로그래밍, 컴퓨터공학, 스마트팩토리, 빅데이터, 디지털헬스")
        ]
    },
    {
        "title": "[스타일-클린학술] 스위스 미니멀 (Swiss-Minimal)",
        "types": ["디자인스타일", "테마", "클린학술"],
        "blocks": [
            callout_block("엄격한 그리드 시스템과 절제된 타이포그래피로 신뢰성과 학술적 권위를 극대화한 미니멀리즘 스타일입니다.", "🏛️"),
            h2_block("1. 디자인 특성 및 색상 가이드"),
            bullet_block(" 장식을 완전히 배제한 샤프한 직각 그리드와 과감하고 정갈한 여백.", "쉐입(Shape):"),
            bullet_block(" 순백색 또는 미니멀 페이퍼 화이트 (#FFFFFF, #F9F9FB)", "캔버스 배경:"),
            bullet_block(" 짙은 먹색 차콜 블랙 (#111111)", "텍스트 기본색:"),
            bullet_block(" 딥 슬레이트 그레이 (#334155) 및 스위스 레드 (#E11D48)", "포인트색:"),
            divider_block(),
            h2_block("2. 추천 대상 과목"),
            bullet_block(" 법학, 공공행정학, 철학, 학술 세미나, 이론 정립형 강의 및 논문 리뷰")
        ]
    },
    {
        "title": "[스타일-비즈니스] MBB 컨설팅 (Consulting-Framework)",
        "types": ["디자인스타일", "테마", "비즈니스"],
        "blocks": [
            callout_block("맥킨지/BCG 전략 컨설팅 보고서 형식으로, 슬라이드 상단의 핵심 메시지와 정교한 3~4단 프레임워크 카드가 결합된 고밀도 비즈니스 스타일입니다.", "💼"),
            h2_block("1. 디자인 특성 및 색상 가이드"),
            bullet_block(" 명확한 구획 분할, 2단/3단 비교 카드, 프로세스 셰브론 및 매트릭스 그리드.", "쉐입(Shape):"),
            bullet_block(" 포멀 라이트 그레이 (#F1F5F9, #FFFFFF)", "캔버스 배경:"),
            bullet_block(" 전문적인 딥 차콜 (#0F172A)", "텍스트 기본색:"),
            bullet_block(" 코퍼레이트 네이비 블루 (#1E3A8A)", "주색상(Primary):"),
            bullet_block(" 컨설팅 골드/앰버 (#D97706) 또는 로열 블루 (#2563EB)", "포인트색(Accent):"),
            divider_block(),
            h2_block("2. 추천 대상 과목"),
            bullet_block(" 경영전략, 글로벌AI커머스, 마케팅, 경제학, 재무회계, 창업 및 벤처")
        ]
    },
    {
        "title": "[스타일-강의실] 클래식 칠판 (Chalkboard)",
        "types": ["디자인스타일", "테마", "강의실"],
        "blocks": [
            callout_block("전통적인 대학 강의실 칠판 느낌의 다크 그린 캔버스 위에 분필 손글씨/판서 도해 느낌을 구현한 직관적 강의 스타일입니다.", "👨‍🏫"),
            h2_block("1. 디자인 특성 및 색상 가이드"),
            bullet_block(" 정형화되지 않은 친근한 박스 테두리선 및 분필 텍스처 느낌의 도해 도형.", "쉐입(Shape):"),
            bullet_block(" 다크 슬레이트 그린 칠판색 (#1A2E26 또는 #1F2421)", "캔버스 배경:"),
            bullet_block(" 부드러운 초크 화이트 (#F8FAFC)", "텍스트 기본색:"),
            bullet_block(" 분필 옐로우 (#FDE047) 및 파스텔 민트 (#6EE7B7)", "분필 포인트:"),
            bullet_block(" 파스텔 피치 핑크 (#FDA4AF)", "강조색:"),
            divider_block(),
            h2_block("2. 추천 대상 과목"),
            bullet_block(" 수학, 통계학, 기초자연과학, 논리적 수식 유도, 아이디어 발상 및 워크숍")
        ]
    }
]

def seed():
    print("=" * 70)
    print(" [Notion PPT-Guidelines] 대학 강의 5대 핵심 비주얼 스타일 등록 시작")
    print("=" * 70)
    for s in STYLES_DATA:
        title = s["title"]
        types = s["types"]
        blocks = s["blocks"]
        
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

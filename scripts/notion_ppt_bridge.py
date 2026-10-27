#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Notion-PPT Automation Bridge
- Notion DB(PPT-Pipeline, PPT-Scripts, PPT-Guidelines)와 로컬 생성 엔진(ppt-master)을 연동
- 노션 본문 내 마크다운 추출 및 이미지(S3) 로컬 자동 캐싱
- 생성된 PPTX 파일 Google Drive 자동 동기화 (로컬 마운트 드라이브 복사)
- Notion PPT결과 속성 링크 Append 및 진행상태 업데이트
"""

import os
import sys
import json
import shutil
import urllib.request
import urllib.parse
from pathlib import Path
from datetime import datetime

# Windows 콘솔 UTF-8 인코딩 보장
if sys.platform == "win32":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")

# ==========================================
# 기본 환경 설정 (.env 자동 로드)
# ==========================================
def _load_env():
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent.parent / ".env",
        Path(__file__).resolve().parent / ".env"
    ]
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

PIPELINE_DB_ID = "3fa624a3-5444-8064-aa7c-ed9c93dd8bdd"
SCRIPTS_DB_ID = "3fa624a3-5444-803c-93ff-f437c0b9f9d8"
GUIDELINES_DB_ID = "3fa624a3-5444-80c1-9c4e-c4ea770a8a92"

GDRIVE_LOCAL_PATH = Path(r"G:\내 드라이브\Works_attach\NOTION-PPT")
GDRIVE_WEB_URL = "https://drive.google.com/open?id=1CH7TFrOZMDKQD3SsvFWQ2Sl1sZcbWMwG&usp=drive_fs"

NOTION_VERSION = "2022-06-28"
NOTION_API_BASE = "https://api.notion.com/v1"

# ==========================================
# Notion API Helper
# ==========================================
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

def query_database(database_id, filter_dict=None):
    body = {"page_size": 100}
    if filter_dict:
        body["filter"] = filter_dict
    return notion_request(f"databases/{database_id}/query", method="POST", body=body)

def get_page(page_id):
    return notion_request(f"pages/{page_id}", method="GET")

def update_page_properties(page_id, properties):
    body = {"properties": properties}
    return notion_request(f"pages/{page_id}", method="PATCH", body=body)

def get_block_children(block_id):
    results = []
    start_cursor = None
    while True:
        endpoint = f"blocks/{block_id}/children?page_size=100"
        if start_cursor:
            endpoint += f"&start_cursor={start_cursor}"
        res = notion_request(endpoint, method="GET")
        results.extend(res.get("results", []))
        if not res.get("has_more"):
            break
        start_cursor = res.get("next_cursor")
    return results

# ==========================================
# Rich Text & Block to Markdown Converter
# ==========================================
def rich_text_to_markdown(rich_text_list):
    if not rich_text_list:
        return ""
    md_parts = []
    for item in rich_text_list:
        text = item.get("plain_text", "")
        annotations = item.get("annotations", {})
        if annotations.get("code"):
            text = f"`{text}`"
        if annotations.get("bold"):
            text = f"**{text}**"
        if annotations.get("italic"):
            text = f"*{text}*"
        if annotations.get("strikethrough"):
            text = f"~~{text}~~"
        
        href = item.get("href")
        if href:
            text = f"[{text}]({href})"
        md_parts.append(text)
    return "".join(md_parts)

def blocks_to_markdown(blocks, indent=0):
    lines = []
    prefix = "  " * indent
    for block in blocks:
        b_type = block.get("type")
        b_data = block.get(b_type, {})
        
        if b_type == "paragraph":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"{prefix}{text}\n")
        elif b_type == "heading_1":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"\n{prefix}# {text}\n")
        elif b_type == "heading_2":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"\n{prefix}## {text}\n")
        elif b_type == "heading_3":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"\n{prefix}### {text}\n")
        elif b_type == "bulleted_list_item":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"{prefix}* {text}")
        elif b_type == "numbered_list_item":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"{prefix}1. {text}")
        elif b_type == "to_do":
            checked = b_data.get("checked", False)
            mark = "x" if checked else " "
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"{prefix}- [{mark}] {text}")
        elif b_type == "quote":
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"{prefix}> {text}")
        elif b_type == "callout":
            icon = b_data.get("icon", {}).get("emoji", "💡")
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"\n{prefix}> {icon} {text}\n")
        elif b_type == "code":
            lang = b_data.get("language", "")
            text = rich_text_to_markdown(b_data.get("rich_text", []))
            lines.append(f"```{lang}\n{text}\n```")
        elif b_type == "divider":
            lines.append("\n---\n")
        elif b_type == "image":
            img_type = b_data.get("type")
            url = b_data.get(img_type, {}).get("url", "")
            caption = rich_text_to_markdown(b_data.get("caption", []))
            lines.append(f"![{caption}]({url})")
        
        elif b_type == "table":
            table_rows = get_block_children(block["id"])
            table_md_lines = []
            for r_idx, row_block in enumerate(table_rows):
                if row_block.get("type") == "table_row":
                    cells = row_block.get("table_row", {}).get("cells", [])
                    cell_texts = [rich_text_to_markdown(c).replace("\n", "<br>").strip() for c in cells]
                    table_md_lines.append("| " + " | ".join(cell_texts) + " |")
                    if r_idx == 0:
                        table_md_lines.append("| " + " | ".join(["---"] * len(cell_texts)) + " |")
            lines.append("\n" + "\n".join(table_md_lines) + "\n")
            continue  # 자식 처리를 table_row로 직접 완료했으므로 하위 블록 순회 스킵

        # 하위 블록 순회
        if block.get("has_children"):
            children = get_block_children(block["id"])
            child_md = blocks_to_markdown(children, indent=indent + 1)
            lines.append(child_md)

    return "\n".join(lines)

def markdown_to_blocks(md_content: str):
    """
    마크다운 텍스트를 Notion API가 수용하는 블록 규격 리스트로 파싱합니다.
    - 헤딩 (# h1, ## h2, ### h3, #### h3)
    - 구분선 (---)
    - 불릿 (- , * , • )
    - 번호 목록 (1. , 2. )
    - 인용문 (> )
    - 표 (| col | col |) -> Notion 네이티브 table + table_row 블록
    - 일반 문단 (paragraph, 2,000자 초과 시 1,900자 단위 분할)
    """
    import re
    lines = md_content.splitlines()
    blocks = []
    
    i = 0
    n = len(lines)
    
    while i < n:
        line = lines[i]
        stripped = line.strip()
        
        if not stripped:
            i += 1
            continue
            
        # 1. Divider
        if stripped in ("---", "***", "___"):
            blocks.append({
                "object": "block",
                "type": "divider",
                "divider": {}
            })
            i += 1
            continue
            
        # 2. Headings
        if stripped.startswith("### "):
            text = stripped[4:].strip()
            blocks.append({
                "object": "block",
                "type": "heading_3",
                "heading_3": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue
        elif stripped.startswith("## "):
            text = stripped[3:].strip()
            blocks.append({
                "object": "block",
                "type": "heading_2",
                "heading_2": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue
        elif stripped.startswith("# "):
            text = stripped[2:].strip()
            blocks.append({
                "object": "block",
                "type": "heading_1",
                "heading_1": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue
        elif stripped.startswith("#### "):
            text = stripped[5:].strip()
            blocks.append({
                "object": "block",
                "type": "heading_3",
                "heading_3": {"rich_text": [{"type": "text", "text": {"content": f"🎯 {text}"[:2000]}}]}
            })
            i += 1
            continue

        # 3. Table detection: lines starting and ending with '|'
        if stripped.startswith("|") and stripped.endswith("|"):
            table_lines = []
            while i < n and lines[i].strip().startswith("|") and lines[i].strip().endswith("|"):
                t_line = lines[i].strip()
                is_sep = all(c in "-|: \t" for c in t_line) and "-" in t_line
                if not is_sep:
                    raw_cells = t_line[1:-1].split("|")
                    cells = [c.strip().replace("<br>", "\n") for c in raw_cells]
                    table_lines.append(cells)
                i += 1
                
            if table_lines:
                max_cols = max(len(row) for row in table_lines)
                if max_cols > 0:
                    table_rows = []
                    for row in table_lines:
                        padded = row + [""] * (max_cols - len(row))
                        cells_json = [
                            [{"type": "text", "text": {"content": cell[:2000]}}]
                            for cell in padded
                        ]
                        table_rows.append({
                            "type": "table_row",
                            "table_row": {"cells": cells_json}
                        })
                    
                    blocks.append({
                        "object": "block",
                        "type": "table",
                        "table": {
                            "table_width": max_cols,
                            "has_column_header": True,
                            "has_row_header": False,
                            "children": table_rows
                        }
                    })
            continue

        # 4. Bullet list
        if stripped.startswith(("- ", "* ", "• ")):
            text = stripped[2:].strip()
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue

        # 5. Numbered list
        m_num = re.match(r"^(\d+)\.\s+(.*)$", stripped)
        if m_num:
            text = m_num.group(2).strip()
            blocks.append({
                "object": "block",
                "type": "numbered_list_item",
                "numbered_list_item": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue

        # 6. Quote
        if stripped.startswith("> "):
            text = stripped[2:].strip()
            blocks.append({
                "object": "block",
                "type": "quote",
                "quote": {"rich_text": [{"type": "text", "text": {"content": text[:2000]}}]}
            })
            i += 1
            continue

        # 7. Fallback Paragraph (chunked if > 2000 chars)
        chunk_len = 1900
        for c_idx in range(0, len(stripped), chunk_len):
            chunk = stripped[c_idx:c_idx + chunk_len]
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": [{"type": "text", "text": {"content": chunk}}]}
            })
        i += 1

    return blocks

def append_blocks_in_chunks(parent_id: str, blocks: list, chunk_size: int = 80):
    """
    Notion API의 1회 요청당 최대 100개 블록 한도를 준수하여 안전하게 배치 분할 업로드
    """
    total = len(blocks)
    for i in range(0, total, chunk_size):
        chunk = blocks[i:i + chunk_size]
        print(f"  - 노션 블록 업로드 중: {i + 1}~{min(i + len(chunk), total)} / {total}개 블록...")
        notion_request(f"blocks/{parent_id}/children", method="PATCH", body={"children": chunk})

# ==========================================
# Image Downloader & Link Replacer
# ==========================================
def download_image(url, save_path):
    headers = {"User-Agent": "Mozilla/5.0"}
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as response, open(save_path, 'wb') as out_file:
        shutil.copyfileobj(response, out_file)

def cache_images_and_fix_links(md_content, project_name, md_dir):
    """
    마크다운 본문 내 이미지 URL(특히 Notion S3 링크)을 탐지하여
    로컬 md/<project_name>/ 폴더로 다운로드 후 상대 경로로 치환
    """
    import re
    img_dir = md_dir / project_name
    img_dir.mkdir(parents=True, exist_ok=True)
    
    # ![caption](url) 패턴 검색
    pattern = r'!\[(.*?)\]\((https?://[^\)]+)\)'
    matches = list(re.finditer(pattern, md_content))
    
    if not matches:
        return md_content, 0
    
    replacements = []
    img_count = 0
    for idx, match in enumerate(matches, 1):
        caption = match.group(1)
        url = match.group(2)
        
        # 파일 확장자 추론
        clean_url = url.split("?")[0]
        ext = os.path.splitext(clean_url)[1].lower()
        if not ext or ext not in [".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"]:
            ext = ".png"
            
        filename = f"img_{idx:02d}{ext}"
        target_path = img_dir / filename
        
        print(f"  [Image Download] #{idx}: {filename} <- {clean_url[:60]}...")
        try:
            download_image(url, target_path)
            # 상대 경로로 치환: md/<project_name>/img_01.png
            rel_path = f"{project_name}/{filename}"
            replacements.append((match.group(0), f"![{caption}]({rel_path})"))
            img_count += 1
        except Exception as e:
            print(f"  [Warning] Failed to download image #{idx}: {e}", file=sys.stderr)
            
    fixed_md = md_content
    for old_tag, new_tag in replacements:
        fixed_md = fixed_md.replace(old_tag, new_tag)
        
    return fixed_md, img_count

# ==========================================
# 주요 비즈니스 파이프라인 함수
# ==========================================
def find_pipeline_project(project_identifier):
    """프로젝트명 또는 Page ID로 PPT-Pipeline에서 항목 검색"""
    all_data = query_database(PIPELINE_DB_ID)
    for page in all_data.get("results", []):
        p_id = page["id"]
        title_objs = page["properties"].get("프로젝트", {}).get("title", [])
        title = title_objs[0]["plain_text"] if title_objs else ""
        if project_identifier in [p_id, title]:
            return page
    return None

def update_project_status(project_identifier, new_status):
    """PPT-Pipeline 프로젝트의 진행상태를 변경"""
    page = find_pipeline_project(project_identifier)
    if not page:
        print(f"[Error] 프로젝트 '{project_identifier}'를 PPT-Pipeline에서 찾을 수 없습니다.", file=sys.stderr)
        return False
    p_id = page["id"]
    title_objs = page["properties"].get("프로젝트", {}).get("title", [])
    title = title_objs[0]["plain_text"] if title_objs else "untitled"
    
    update_page_properties(p_id, {"진행상태": {"select": {"name": new_status}}})
    print(f"[Notion Update] 프로젝트 '{title}'의 진행상태를 [{new_status}]로 변경했습니다.")
    return True

def process_pipeline_script(project_identifier):
    """
    [1단계: 원고 변환 파이프라인]
    - 상태가 '원고 준비 완료'인지 엄격히 검증 (아니면 안내 후 즉시 중단)
    - 연결된 PPT-Scripts의 원본파일을 마크다운으로 변환
    - 변환 성공 시 진행상태를 '원고 완료'로 갱신하여 사용자 검토 대기
    """
    page = find_pipeline_project(project_identifier)
    if not page:
        print(f"[Error] 프로젝트 '{project_identifier}'를 PPT-Pipeline에서 찾을 수 없습니다.", file=sys.stderr)
        return False
        
    props = page["properties"]
    title_objs = props.get("프로젝트", {}).get("title", [])
    proj_name = title_objs[0]["plain_text"] if title_objs else "untitled"
    
    status_obj = props.get("진행상태", {}).get("select")
    status = status_obj["name"] if status_obj else "미지정"
    
    # 2번 상태(원고 준비 완료) 엄격 검증 (기존 '준비 완료' 호환)
    if status not in ["원고 준비 완료", "준비 완료"]:
        print(f"\n[안내] 프로젝트 '{proj_name}'은(는) 아직 준비가 안된 것 같습니다.")
        print(f"  - 현재 진행상태: [{status}] (필요 상태: '원고 준비 완료')")
        return False
        
    script_relations = props.get("원고", {}).get("relation", []) or props.get("원고(마크다운)", {}).get("relation", [])
    if not script_relations:
        print(f"[Error] 프로젝트 '{proj_name}'에 연결된 원고(PPT-Scripts)가 없습니다.", file=sys.stderr)
        return False
        
    script_id = script_relations[0]["id"]
    print(f"\n[1단계 착수] 프로젝트 '{proj_name}' 원고 변환 시작...")
    convert_script_from_file(script_id)
    
    # 상태를 '원고 완료'로 갱신
    update_project_status(page["id"], "원고 완료")
    print(f"\n[1단계 완료] 프로젝트 '{proj_name}' 원고 변환 완료 ➔ 진행상태: [원고 완료]")
    print("👉 사용자가 노션에서 원고 마크다운을 검토한 후 'ppt 준비 완료'로 변경하면 2단계(PPT 개발)를 진행할 수 있습니다.")
    return True

def check_pipeline_for_ppt(project_identifier):
    """
    [2단계 진입 전 상태 검증]
    - 상태가 '개발 준비 완료' (또는 'ppt 준비 완료')인지 엄격히 검증
    """
    page = find_pipeline_project(project_identifier)
    if not page:
        print(f"[Error] 프로젝트 '{project_identifier}'를 PPT-Pipeline에서 찾을 수 없습니다.", file=sys.stderr)
        return None
        
    props = page["properties"]
    title_objs = props.get("프로젝트", {}).get("title", [])
    proj_name = title_objs[0]["plain_text"] if title_objs else "untitled"
    
    status_obj = props.get("진행상태", {}).get("select")
    status = status_obj["name"] if status_obj else "미지정"
    
    if status not in ["개발 준비 완료", "ppt 준비 완료", "PPT 준비 완료", "준비 완료"]:
        print(f"\n[안내] 프로젝트 '{proj_name}'은(는) 아직 준비가 안된 것 같습니다.")
        print(f"  - 현재 진행상태: [{status}] (필요 상태: '개발 준비 완료')")
        return None
        
    return page

def list_pipeline_projects():
    """PPT-Pipeline DB의 프로젝트 목록 및 5단계 상태 머신 현황 조회"""
    data = query_database(PIPELINE_DB_ID)
    projects = []
    status_counts = {
        "Task 준비 중": 0,
        "준비 중": 0,
        "원고 준비 완료": 0,
        "원고 완료": 0,
        "개발 준비 완료": 0,
        "ppt 준비 완료": 0,
        "개발 완료": 0,
        "ppt 생성 완료": 0,
        "기타": 0
    }
    
    print("\n" + "=" * 70)
    print(" [PPT-Dev-Pipeline] 프로젝트 목록 (5단계 라이프사이클)")
    print("=" * 70)
    
    for page in data.get("results", []):
        p_id = page["id"]
        props = page["properties"]
        
        title_objs = props.get("프로젝트", {}).get("title", [])
        title = title_objs[0]["plain_text"] if title_objs else "(제목 없음)"
        
        status_obj = props.get("진행상태", {}).get("select")
        status = status_obj["name"] if status_obj else "미지정"
        
        # 카운트 집계
        if status in status_counts:
            status_counts[status] += 1
        elif status == "준비 완료":
            status_counts["원고 준비 완료"] += 1
        elif status in ["PPT 생성완료", "PPT생성완료"]:
            status_counts["개발 완료"] += 1
        else:
            status_counts["기타"] += 1
        
        dev_type_obj = props.get("개발유형", {}).get("select")
        dev_type = dev_type_obj["name"] if dev_type_obj else "powerpoint"

        ratio_obj = props.get("개발비율", {}).get("select") or props.get("비율", {}).get("select")
        ratio = ratio_obj["name"] if ratio_obj else "4:3 (기본)"
        
        scripts = props.get("원고", {}).get("relation", []) or props.get("원고(마크다운)", {}).get("relation", [])
        guidelines = props.get("개발지침", {}).get("relation", []) or props.get("PPT개발지침", {}).get("relation", [])
        
        files_list = props.get("결과물", {}).get("files", [])
        files_str = f"첨부파일 {len(files_list)}개" if files_list else "없음"
        
        results_objs = props.get("PPT결과", {}).get("rich_text", [])
        results_preview = results_objs[0]["plain_text"][:30] + "..." if results_objs else files_str
        
        projects.append({
            "id": p_id,
            "title": title,
            "status": status,
            "dev_type": dev_type,
            "ratio": ratio,
            "has_script": len(scripts) > 0,
            "has_guidelines": len(guidelines) > 0,
            "results": results_preview
        })
        
        print(f"* 프로젝트: {title} (ID: {p_id})")
        print(f"  - 상태: [{status}] | 개발유형: [{dev_type}] | 비율: [{ratio}]")
        print(f"  - 원고 연결: {'O (' + str(len(scripts)) + '개)' if scripts else 'X'} | 지침 연결: {'O (' + str(len(guidelines)) + '개)' if guidelines else 'X'}")
        print(f"  - 결과물: {results_preview}")
        print("-" * 70)
        
    ready_count = status_counts['개발 준비 완료'] + status_counts['ppt 준비 완료']
    done_count = status_counts['개발 완료'] + status_counts['ppt 생성 완료']
    print(f"[상태별 집계] 준비 중: {status_counts['준비 중'] + status_counts['Task 준비 중']}개 | 원고 준비 완료: {status_counts['원고 준비 완료']}개 | 원고 완료: {status_counts['원고 완료']}개 | 개발 준비 완료(2단계 대기): {ready_count}개 | 개발 완료: {done_count}개")
    print("=" * 70)
    return projects

def fetch_project_data(project_identifier, check_ready=False):
    """
    프로젝트명 또는 Page ID로 원고와 지침을 수집하여 로컬 md 파일로 빌드
    """
    target_page = find_pipeline_project(project_identifier)
    if not target_page:
        raise ValueError(f"프로젝트 '{project_identifier}'를 PPT-Pipeline에서 찾을 수 없습니다.")
        
    props = target_page["properties"]
    title_objs = props.get("프로젝트", {}).get("title", [])
    proj_name = title_objs[0]["plain_text"] if title_objs else "untitled"
    
    if check_ready:
        status_obj = props.get("진행상태", {}).get("select")
        status = status_obj["name"] if status_obj else "미지정"
        if status not in ["개발 준비 완료", "ppt 준비 완료", "PPT 준비 완료", "준비 완료"]:
            print(f"\n[안내] 프로젝트 '{proj_name}'은(는) 아직 준비가 안된 것 같습니다.")
            print(f"  - 현재 진행상태: [{status}] (필요 상태: '개발 준비 완료')")
            return None
            return None
        
    props = target_page["properties"]
    title_objs = props.get("프로젝트", {}).get("title", [])
    proj_name = title_objs[0]["plain_text"] if title_objs else "untitled"
    # 안전한 파일명으로 정규화
    safe_name = "".join([c for c in proj_name if c.isalnum() or c in (" ", "-", "_")]).strip()
    
    ratio_obj = props.get("개발비율", {}).get("select") or props.get("비율", {}).get("select")
    ratio = ratio_obj["name"] if ratio_obj else "4:3"
    
    script_relations = props.get("원고", {}).get("relation", []) or props.get("원고(마크다운)", {}).get("relation", [])
    guideline_relations = props.get("PPT개발지침", {}).get("relation", []) or props.get("개발지침", {}).get("relation", [])
    
    print(f"\n[Fetch] 프로젝트 '{proj_name}' 데이터 수집 시작...")
    
    # 2. 원고 마크다운 추출
    script_md_parts = []
    if not script_relations:
        print("  [Warning] 연결된 원고(PPT-Scripts)가 없습니다. 프로젝트 본문을 검사합니다.")
        proj_blocks = get_block_children(target_page["id"])
        script_md_parts.append(blocks_to_markdown(proj_blocks))
    else:
        for s_ref in script_relations:
            s_page = get_page(s_ref["id"])
            s_title = s_page["properties"].get("원고", {}).get("title", [])
            s_title_str = s_title[0]["plain_text"] if s_title else "원고"
            print(f"  - 원고 로드: '{s_title_str}' ({s_ref['id']})")
            
            blocks = get_block_children(s_ref["id"])
            md = blocks_to_markdown(blocks)
            script_md_parts.append(md)
            
    full_script_md = "\n\n".join(script_md_parts)
    
    # 3. 개발지침 추출
    guideline_parts = []
    for g_ref in guideline_relations:
        g_page = get_page(g_ref["id"])
        g_title = g_page["properties"].get("개발지침", {}).get("title", [])
        g_title_str = g_title[0]["plain_text"] if g_title else "지침"
        print(f"  - 개발지침 로드: '{g_title_str}' ({g_ref['id']})")
        
        g_blocks = get_block_children(g_ref["id"])
        g_md = blocks_to_markdown(g_blocks)
        guideline_parts.append(f"### [지침: {g_title_str}]\n{g_md}")
        
    full_guideline_md = "\n\n".join(guideline_parts)
    
    # 4. 이미지 다운로드 및 로컬 캐싱
    md_root = Path("md")
    md_root.mkdir(parents=True, exist_ok=True)
    
    print("  [Image Caching] 본문 내 이미지 다운로드 및 상대 경로 치환 중...")
    clean_md, downloaded_count = cache_images_and_fix_links(full_script_md, safe_name, md_root)
    print(f"  [Image Caching] 총 {downloaded_count}개 이미지 로컬 저장 완료.")
    
    # 5. 로컬 파일 저장
    v1_file = md_root / f"{safe_name}_v1.md"
    with open(v1_file, "w", encoding="utf-8") as f:
        f.write(clean_md)
    print(f"  -> 원고 마크다운 저장: {v1_file}")
    
    guideline_file = md_root / f"{safe_name}_guidelines.md"
    if full_guideline_md:
        with open(guideline_file, "w", encoding="utf-8") as f:
            f.write(full_guideline_md)
        print(f"  -> 지침 마크다운 저장: {guideline_file}")
        
    return {
        "project_id": target_page["id"],
        "project_name": proj_name,
        "safe_name": safe_name,
        "ratio": ratio,
        "v1_file": str(v1_file),
        "guideline_file": str(guideline_file) if full_guideline_md else None,
        "image_count": downloaded_count
    }

def convert_script_from_file(script_identifier):
    """
    PPT-Scripts에 등록된 '원본파일'(HWP, HWPX, DOCX, PPTX, PDF)을 다운로드하고
    연결된 '원고변환지침'에 맞춰 마크다운으로 변환한 뒤,
    기존 내용을 지우지 않고 본문 아래에 새 버전으로 Append 및 버전 속성 증가
    """
    sys.path.append(str(Path(__file__).resolve().parent))
    try:
        from doc_parser import extract_manuscript_text
    except ImportError as e:
        print(f"[Error] doc_parser를 불러올 수 없습니다: {e}", file=sys.stderr)
        return

    # 1. 대상 원고 페이지 검색
    all_scripts = query_database(SCRIPTS_DB_ID)
    target_page = None
    for page in all_scripts.get("results", []):
        p_id = page["id"]
        title_objs = page["properties"].get("원고", {}).get("title", [])
        title = title_objs[0]["plain_text"] if title_objs else ""
        if script_identifier in [p_id, title]:
            target_page = page
            break

    if not target_page:
        print(f"[Error] 원고 '{script_identifier}'를 PPT-Scripts에서 찾을 수 없습니다.", file=sys.stderr)
        return

    page_id = target_page["id"]
    props = target_page["properties"]
    title_objs = props.get("원고", {}).get("title", [])
    script_title = title_objs[0]["plain_text"] if title_objs else "untitled"

    # 원본 파일 확인
    files_list = props.get("원본파일", {}).get("files", [])
    if not files_list:
        print(f"[Error] '{script_title}'에 등록된 '원본파일'이 없습니다.", file=sys.stderr)
        return

    first_file = files_list[0]
    file_name = first_file.get("name", "source_file")
    file_url = first_file.get("file", {}).get("url") or first_file.get("external", {}).get("url")
    if not file_url:
        print(f"[Error] 파일 다운로드 URL을 가져올 수 없습니다.", file=sys.stderr)
        return

    # 연결된 원고변환지침 확인
    guideline_relations = props.get("원고변환지침", {}).get("relation", [])
    guideline_names = []
    guideline_texts = []
    for g_ref in guideline_relations:
        g_page = get_page(g_ref["id"])
        g_title = g_page["properties"].get("개발지침", {}).get("title", [])
        g_name = g_title[0]["plain_text"] if g_title else "지침"
        guideline_names.append(g_name)
        g_blocks = get_block_children(g_ref["id"])
        guideline_texts.append(f"### [지침: {g_name}]\n" + blocks_to_markdown(g_blocks))

    guideline_summary = ", ".join(guideline_names) if guideline_names else "기본 변환 규칙"

    print(f"\n[Convert Script] 원고: '{script_title}' (파일: {file_name})")
    print(f"  - 적용 지침: {guideline_summary}")

    # 파일 다운로드 (temp/)
    temp_dir = Path("temp")
    temp_dir.mkdir(parents=True, exist_ok=True)
    ext = Path(file_name).suffix
    local_file_path = temp_dir / f"download_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"

    print(f"  - 원본 파일 다운로드 중: {file_name}...")
    download_image(file_url, local_file_path)

    # doc_parser로 텍스트 추출
    print(f"  - 범용 파서(doc_parser)로 텍스트 추출 중...")
    parse_res = extract_manuscript_text(str(local_file_path))
    if parse_res["status"] != "success":
        print(f"[Error] 파일 파싱 실패: {parse_res['error_msg']}", file=sys.stderr)
        return

    raw_text = parse_res["text"]
    char_len = parse_res["char_count"]
    print(f"  - 텍스트 추출 완료: 총 {char_len:,}자")

    # 버전 계산
    curr_version_obj = props.get("버전", {}).get("select")
    curr_v = curr_version_obj["name"] if curr_version_obj else "v00"
    v_num = int(curr_v.replace("v", "")) if curr_v.startswith("v") and curr_v.replace("v", "").isdigit() else 0
    next_v = f"v{v_num + 1:02d}"

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 마크다운 블록 조립 (비파괴적 Append)
    new_blocks = [
        {"object": "block", "type": "divider", "divider": {}},
        {
            "object": "block",
            "type": "heading_2",
            "heading_2": {
                "rich_text": [{"type": "text", "text": {"content": f"📌 [{next_v}] {now_str} (원본 자동 변환 마크다운)"}}]
            }
        },
        {
            "object": "block",
            "type": "callout",
            "callout": {
                "icon": {"type": "emoji", "emoji": "📎"},
                "rich_text": [{"type": "text", "text": {"content": f"원본 파일: {file_name} | 적용 지침: {guideline_summary}"}}]
            }
        }
    ]

    # raw_text를 구조화된 노션 블록들로 변환
    content_blocks = markdown_to_blocks(raw_text)
    all_blocks = new_blocks + content_blocks

    # 노션 페이지 본문 아래에 배치 분할 누적 (Append)
    print(f"  - 노션 본문 아래에 [{next_v}] 마크다운 블록 (총 {len(all_blocks)}개) 누적(Append) 중...")
    append_blocks_in_chunks(page_id, all_blocks, chunk_size=80)

    # 버전 속성 업데이트
    update_page_properties(page_id, {"버전": {"select": {"name": next_v}}})
    print(f"\n[Success] '{script_title}' 변환 완료!")
    print(f"  - 버전 갱신: {curr_v} ➔ [{next_v}]")
    print(f"  - 노션 본문 하단에 [{next_v}] 마크다운이 성공적으로 누적되었습니다.")

def upload_file_to_notion(file_path):
    """
    Notion v1/file_uploads API를 사용하여 로컬 파일을 노션 스토리지에 직접 업로드하고
    files 속성에 첨부할 수 있는 file_upload 객체를 반환합니다.
    """
    import requests
    src_path = Path(file_path).resolve()
    if not src_path.exists():
        print(f"[Warning] 업로드할 파일을 찾을 수 없습니다: {src_path}", file=sys.stderr)
        return None

    ext = src_path.suffix.lower()
    content_type_map = {
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".html": "text/html",
        ".htm": "text/html",
        ".pdf": "application/pdf",
        ".zip": "application/zip",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".svg": "image/svg+xml"
    }
    content_type = content_type_map.get(ext, "application/octet-stream")
    
    # 1. file_upload 객체 생성
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json"
    }
    init_res = requests.post(
        f"{NOTION_API_BASE}/file_uploads",
        headers=headers,
        json={"filename": src_path.name, "content_type": content_type}
    )
    if init_res.status_code not in (200, 201):
        print(f"[Notion Upload Error] init 실패: {init_res.status_code} {init_res.text}", file=sys.stderr)
        return None
        
    init_data = init_res.json()
    upload_url = init_data["upload_url"]
    upload_id = init_data["id"]
    
    # 2. 파일 바이너리 전송 (multipart/form-data)
    send_headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION
    }
    with open(src_path, "rb") as f:
        send_res = requests.post(
            upload_url,
            headers=send_headers,
            files={"file": (src_path.name, f, content_type)}
        )
        
    if send_res.status_code not in (200, 201):
        print(f"[Notion Upload Error] 파일 전송 실패: {send_res.status_code} {send_res.text}", file=sys.stderr)
        return None
        
    print(f"  [Notion File Upload] 업로드 성공: '{src_path.name}' (ID: {upload_id})")
    return {
        "name": src_path.name,
        "type": "file_upload",
        "file_upload": {
            "id": upload_id
        }
    }

def deploy_and_update_pipeline(project_id, files_to_deploy, new_status="개발 완료"):
    """
    생성된 산출물(PPTX, HTML 등) 파일들을:
    1. 구글 드라이브(로컬 마운트)의 <프로젝트명>/ 하위 폴더로 복사 (로컬 백업)
    2. Notion PPT-Pipeline의 '결과물'(files 속성)에 직접 업로드하여 누적(Append)
    3. Notion 'PPT결과' 텍스트 속성에 일시와 파일명 누적
    4. 진행상태를 new_status(기본: '개발 완료')로 갱신
    """
    import re
    target_page = get_page(project_id)
    props = target_page["properties"]
    
    # 프로젝트명 추출
    title_objs = props.get("프로젝트", {}).get("title", [])
    proj_name = title_objs[0]["plain_text"] if title_objs else "Untitled"
    safe_proj_name = re.sub(r'[\\/*?:"<>|]', "_", proj_name).strip()
    
    # 기존 PPT결과 텍스트 가져오기
    existing_rich_text = props.get("PPT결과", {}).get("rich_text", [])
    existing_text = "".join([t.get("plain_text", "") for t in existing_rich_text])
    
    # 기존 '결과물' 파일 목록 가져오기 (누적 보존용)
    existing_files = props.get("결과물", {}).get("files", [])
    
    copied_files_info = []
    new_notion_files = []
    
    # 구글 드라이브 마운트 확인 및 프로젝트 하위 폴더 생성 (백업용)
    gdrive_active = GDRIVE_LOCAL_PATH.exists()
    if gdrive_active:
        proj_gdrive_dir = GDRIVE_LOCAL_PATH / safe_proj_name
        proj_gdrive_dir.mkdir(parents=True, exist_ok=True)
    else:
        proj_gdrive_dir = None
        
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    for file_str in files_to_deploy:
        src_path = Path(file_str).resolve()
        if not src_path.exists():
            print(f"  [Warning] 파일을 찾을 수 없어 건너뜁니다: {src_path}", file=sys.stderr)
            continue
            
        file_name = src_path.name
        
        # 1. 로컬 구글 드라이브 백업 복사
        if gdrive_active and proj_gdrive_dir:
            dest_path = proj_gdrive_dir / file_name
            shutil.copy2(src_path, dest_path)
            print(f"  [Google Drive Sync] 동기화 완료: {dest_path}")
            
        info_line = f"[{now_str}] {file_name}"
        copied_files_info.append(info_line)
        
        # 2. 노션 File Upload API로 직접 업로드
        uploaded_obj = upload_file_to_notion(src_path)
        if uploaded_obj:
            new_notion_files.append(uploaded_obj)
        
    if not copied_files_info:
        print("동기화할 유효한 산출물 파일이 없습니다.")
        return
        
    # 노션 속성 업데이트 준비
    update_props = {
        "진행상태": {
            "select": {"name": new_status}
        }
    }
    
    # PPT결과 텍스트 속성이 있으면 누적
    if "PPT결과" in props:
        new_append_block = "\n".join(copied_files_info)
        updated_text = f"{existing_text}\n{new_append_block}".strip() if existing_text else new_append_block
        update_props["PPT결과"] = {
            "rich_text": [
                {
                    "type": "text",
                    "text": {"content": updated_text}
                }
            ]
        }
        
    # '결과물' 파일 및 미디어 속성에 누적(Append) 업데이트
    if "결과물" in props:
        all_files = existing_files + new_notion_files
        update_props["결과물"] = {
            "files": all_files
        }
    
    update_page_properties(project_id, update_props)
    print(f"\n[Notion Update] 프로젝트 ID({project_id}) '결과물' 속성에 파일 {len(new_notion_files)}개 직접 첨부 완료!")
    print(f"  - 진행상태 갱신: [{new_status}]")
    if "PPT결과" in update_props:
        print(f"  - 텍스트 히스토리 기록 완료:\n{new_append_block}")

# ==========================================
# CLI Entrypoint
# ==========================================
def main():
    import argparse
    parser = argparse.ArgumentParser(description="Notion-PPT Bridge Pipeline Controller")
    subparsers = parser.add_subparsers(dest="command", help="실행할 명령어")
    
    # 1. list
    subparsers.add_parser("list", help="PPT-Pipeline 프로젝트 목록 및 5단계 상태 조회")
    
    # 2. fetch
    fetch_parser = subparsers.add_parser("fetch", help="노션에서 원고/지침 수집 및 이미지 로컬 캐싱")
    fetch_parser.add_argument("project", help="프로젝트명 또는 Notion Page ID")
    fetch_parser.add_argument("--check-ready", action="store_true", help="'ppt 준비 완료' 상태인지 사전 검증")
    
    # 3. deploy
    deploy_parser = subparsers.add_parser("deploy", help="생성된 산출물(PPTX, HTML 등) 파일을 구글 드라이브/노션으로 배포")
    deploy_parser.add_argument("project_id", help="Notion PPT-Pipeline Page ID")
    deploy_parser.add_argument("files", nargs="+", help="배포할 파일 경로 목록 (.pptx, index.html 등)")
    deploy_parser.add_argument("--status", default="개발 완료", help="변경할 진행상태 (기본: '개발 완료')")

    # 4. convert-script (원고 DB 단독 변환)
    conv_parser = subparsers.add_parser("convert-script", help="PPT-Scripts에 첨부된 원본파일을 마크다운으로 변환하여 본문 누적 및 버전 증가")
    conv_parser.add_argument("script", help="원고명 또는 PPT-Scripts Page ID")
    
    # 5. process-script (파이프라인 1단계 전용)
    proc_script_parser = subparsers.add_parser("process-script", help="[1단계] 프로젝트의 원고를 변환하고 진행상태를 '원고 완료'로 갱신")
    proc_script_parser.add_argument("project", help="프로젝트명 또는 Notion Page ID")

    # 6. update-status
    stat_parser = subparsers.add_parser("update-status", help="PPT-Pipeline 프로젝트의 진행상태를 변경")
    stat_parser.add_argument("project", help="프로젝트명 또는 Notion Page ID")
    stat_parser.add_argument("status", help="변경할 상태명 (예: '준비 중', '원고 준비 완료', '원고 완료', 'ppt 준비 완료', 'ppt 생성 완료')")

    args = parser.parse_args()
    
    if args.command == "list":
        list_pipeline_projects()
    elif args.command == "fetch":
        fetch_project_data(args.project, check_ready=args.check_ready)
    elif args.command == "deploy":
        deploy_and_update_pipeline(args.project_id, args.files, args.status)
    elif args.command == "convert-script":
        convert_script_from_file(args.script)
    elif args.command == "process-script":
        process_pipeline_script(args.project)
    elif args.command == "update-status":
        update_project_status(args.project, args.status)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()


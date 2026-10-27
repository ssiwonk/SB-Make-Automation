# -*- coding: utf-8 -*-
"""
HTML 슬라이드 자동 캡처 도구 (Playwright 기반)
1920x1080 해상도로 각 슬라이드를 순회하며 실제 렌더링 화면을 스크린샷으로 저장합니다.
"""
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

def capture_slides(html_path_str, output_dir_str="temp/screenshots"):
    html_path = Path(html_path_str).resolve()
    if not html_path.exists():
        print(f"Error: File not found: {html_path}", file=sys.stderr)
        return False

    out_dir = Path(output_dir_str).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    file_url = html_path.as_uri()
    print(f"[Capture] Opening: {file_url}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 16:9 1920x1080 고정 해상도
        page = browser.new_page(viewport={"width": 1920, "height": 1080}, device_scale_factor=1)
        page.goto(file_url, wait_until="networkidle")
        time.sleep(1.0) # 웹폰트 및 Chart.js 로딩 대기

        # 슬라이드 총 개수 확인
        slide_count = page.evaluate("() => document.querySelectorAll('.slide').length")
        print(f"[Capture] Total slides detected: {slide_count}")

        saved_files = []
        for i in range(slide_count):
            if i > 0:
                # 다음 슬라이드로 이동 (키보드 ArrowRight)
                page.keyboard.press("ArrowRight")
                time.sleep(0.5) # 전환 애니메이션 및 차트 렌더링 대기

            out_file = out_dir / f"slide_{i+1:02d}.png"
            page.screenshot(path=str(out_file), full_page=False)
            print(f"  - Slide {i+1:02d} captured: {out_file.name}")
            saved_files.append(out_file)

        browser.close()
        print(f"[Capture] Complete! {len(saved_files)} screenshots saved to {out_dir}")
        return saved_files

if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "project/2610271345_콘텐츠팀_업무_보고/index_v5.html"
    capture_slides(target)

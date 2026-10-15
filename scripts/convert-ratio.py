import os
import sys
import argparse
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

# ==========================================
# 고정 설정값
# ==========================================
LEFT_MARGIN = Inches(0.3)  # 좌측으로 이동시킬 때의 고정 여백 (0.5인치)
# ==========================================


def convert_43_to_169_left_aligned(input_path, output_path=None):
    input_p = Path(input_path).resolve()
    if not input_p.exists():
        raise FileNotFoundError(f"Input PPTX file not found: {input_p}")
    
    if output_path is None:
        output_p = input_p.parent / f"{input_p.stem}-169{input_p.suffix}"
    else:
        output_p = Path(output_path).resolve()

    # PPTX 파일 열기
    prs = Presentation(str(input_p))
    
    # 기존 크기 (보통 4:3은 10 x 7.5 인치)
    old_width = prs.slide_width
    old_height = prs.slide_height
    
    # 새 16:9 크기 (13.333 x 7.5 인치)
    new_width = Inches(13.333)
    new_height = Inches(7.5)
    
    # 1. "맞춤 확인(Ensure Fit)"을 위한 배율(Scale) 계산
    # (새 크기와 기존 크기의 비율 중 더 작은 값을 사용하여 잘리지 않게 맞춤)
    scale = min(new_width / old_width, new_height / old_height)
    
    # 2. 모든 슬라이드의 도형을 순회하며 배율 적용 및 좌측 이동
    for slide in prs.slides:
        if not slide.shapes:
            continue
            
        # 해당 슬라이드에서 가장 왼쪽에 있는 도형의 위치 찾기 (기준점)
        left_positions = [shape.left for shape in slide.shapes if shape.left is not None]
        if not left_positions:
            continue
            
        min_left = min(left_positions)
        
        # 각 도형 크기 및 위치 업데이트
        for shape in slide.shapes:
            try:
                # 크기 조정 (배율 적용)
                if shape.width is not None and shape.height is not None:
                    shape.width = int(shape.width * scale)
                    shape.height = int(shape.height * scale)
                
                # 위치 조정
                if shape.left is not None and shape.top is not None:
                    shape.top = int(shape.top * scale)
                    
                    # 좌측 이동 로직: 
                    # 원래 도형간의 간격((shape.left - min_left) * scale)은 유지하면서
                    # 전체 그룹의 시작점을 고정된 LEFT_MARGIN으로 맞춤
                    shape.left = int(LEFT_MARGIN + ((shape.left - min_left) * scale))
            except Exception as e:
                # 표(Table)나 일부 특수 그룹 도형은 직접 크기 수정이 안 될 수 있으므로 예외 처리
                print(f"Warning: Could not resize/move a shape: {e}")

    # 3. 슬라이드 자체의 해상도/크기 변경
    prs.slide_width = new_width
    prs.slide_height = new_height
    
    # 변경된 파일 저장
    prs.save(str(output_p))
    print(f"[Done] Converted 4:3 -> 16:9: {output_p} (Margin: 0.5 inches)")
    return str(output_p)

def main():
    parser = argparse.ArgumentParser(description="Convert PPTX from 4:3 to 16:9 (Ensure Fit & Left Aligned)")
    parser.add_argument("input_file", nargs="?", default="input.pptx", help="Path to input 4:3 PPTX file")
    parser.add_argument("output_file", nargs="?", default=None, help="Path to output 16:9 PPTX file (optional)")
    
    args = parser.parse_args()
    
    if os.path.exists(args.input_file):
        convert_43_to_169_left_aligned(args.input_file, args.output_file)
    else:
        if len(sys.argv) > 1:
            print(f"Error: File not found: {args.input_file}")
        else:
            print("Usage: python convert-ratio.py <input.pptx> [<output.pptx>]")

if __name__ == "__main__":
    main()
# AGENTS.md

이 문서는 이 프로젝트에서 작업하는 모든 AI 코딩 어시스턴트(Claude, Gemini 등)가 반드시 따라야 하는 공통 작업 규칙입니다.
작업을 시작하기 전에 이 문서 전체를 읽고, 아래 규칙을 예외 없이 따르세요.

---

# [단일 표준] 노션 연동 PPTX 자동화 파이프라인 (Notion-to-PPT Pipeline)

**중요 원칙: 본 프로젝트의 모든 PPTX 저작 및 변환 작업은 오직 노션(Notion) 파이프라인을 통해서만 수행됩니다. 로컬 파일 기반 직접 변환 워크플로우는 허용되지 않습니다.**

노션에 작성된 원고와 재활용 가능한 개발 지침을 결합하여 고품질의 PPTX를 자동 생성하고, 구글 드라이브 동기화 및 노션 결과를 실시간 업데이트하는 단일 통합 파이프라인입니다.

## 1. 노션 데이터베이스 아키텍처
- **상위 대시보드 페이지**: `NOTION-PPT` (`3fa624a3-5444-80bc-98f0-e9ba0624d2e2`)
- **1) 메인 관제 DB**: `PPT-Pipeline` (`3fa624a3-5444-8064-aa7c-ed9c93dd8bdd`)
  - 속성: `프로젝트`(Title), `진행상태`(Select), `개발비율`(Select: 4:3 / 16:9), `원고`(Relation to PPT-Scripts), `PPT개발지침`(Relation to PPT-Guidelines), `PPT결과`(Rich Text), `메모`, `최종편집일`
  - `진행상태` 라이프사이클: `준비 중` ➔ `준비 완료` (AI 작업 트리거) ➔ `PPT 생성완료` (AI 완료)
  - *사용자가 대화창에서 직접 "PPT 개발해줘"라고 요청하면 진행상태와 무관하게 즉시 개발을 시작한다.*
- **2) 원고 DB**: `PPT-Scripts` (`3fa624a3-5444-803c-93ff-f437c0b9f9d8`)
  - 속성: `원고`(Title), `원본파일`(Files: HWP/HWPX/DOCX/PPTX/PDF), `원고변환지침`(Relation to PPT-Guidelines), `버전`(Select: v01~v20), `프로젝트명`(Relation to PPT-Pipeline)
  - 본문(Content): 변환된 마크다운이 타임스탬프 헤더와 함께 **지워지지 않고 아래에 누적(Append)**되며 `버전` 속성이 자동 갱신됨
- **3) 개발지침 DB**: `PPT-Guidelines` (`3fa624a3-5444-80c1-9c4e-c4ea770a8a92`)
  - 속성: `개발지침`(Title), `유형`(Multi-select: `SCU 표준`, `원고변환`, `대본형`, `교재형`, `디자인`, `레이아웃` 등)
  - 본문(Content): 모듈화된 원고 변환 지침, 슬라이드 레이아웃(3단 카드, 대조형 등), 톤앤매너, 컬러 및 타이포그래피 규칙

## 2. 구글 드라이브 및 배포 환경
- **구글 드라이브 로컬 동기화 마운트**: `G:\내 드라이브\Works_attach\NOTION-PPT` (하위에 `<프로젝트명>/` 폴더 자동 생성 및 저장)
- **구글 드라이브 웹 URL**: `https://drive.google.com/open?id=1CH7TFrOZMDKQD3SsvFWQ2Sl1sZcbWMwG&usp=drive_fs`
- **배포 규칙**: 생성된 PPTX는 위 로컬 폴더의 `<프로젝트명>/` 하위 폴더로 복사되어 Google Drive 클라우드에 자동 동기화되며, `PPT-Pipeline`의 `PPT결과` 속성에는 `[YYYY-MM-DD HH:MM] <파일명.pptx>` 형태로 일시와 파일명만 깔끔하게 누적(Append) 기재된다. (드라이브 폴더 링크는 중복 방지를 위해 생략)

## 3. 폴더 구조 규칙
- `md/`     : 노션에서 가져온 원고/지침 마크다운 및 이미지 캐시 폴더
              → `md/<프로젝트명>_v1.md`: 노션 본문 수집 초안
              → `md/<프로젝트명>_v2.md`: `spec/md-definition.md`에 따라 정제된 최종 마크다운
              → `md/<프로젝트명>/`: 노션 본문의 S3 이미지를 영구 보관하는 로컬 캐시 폴더
- `project/YYMMDDHHMM_<프로젝트명>/` : pptx 작업 최종 산출물 전용 폴더
              → 폴더 내부에는 (a) 최종 확정 md 사본, (b) ppt-master 하위 프로젝트 폴더가 위치
              → 최종 pptx 실체 파일은 `exports/<project_name>_<timestamp>.pptx` (4:3 기본) 및 `exports/<project_name>_<timestamp>-169.pptx` (16:9 변환본)에 위치
- `spec/`   : 프로젝트 표준 정의서
              → `spec/md-definition.md` : 원고 Markdown 정제 및 교정 규칙 정의서
              → `spec/pptx-definition.md` : Markdown → PPTX 디자인, 인포그래픽 및 SVG 제작 규칙 정의서
- `scripts/`: 파이프라인 자동화 브릿지 및 변환 도구
              → `notion_ppt_bridge.py`: 노션 ↔ 로컬 ↔ 구글드라이브 연동 핵심 브릿지
              → `convert-ratio.py`: 4:3 ➔ 16:9 와이드 비율 변환 스크립트
              → `doc_parser.py`: HWP/HWPX/DOCX/PDF/PPTX 범용 텍스트 추출 모듈
- `temp/`   : AI 임시 작업용 폴더 (언제든 삭제 가능)

---

## 4. 파이프라인 자동화 브릿지 스크립트 (`scripts/notion_ppt_bridge.py`)

- **프로젝트 목록 조회**:
  ```bash
  python scripts/notion_ppt_bridge.py list
  ```
- **원고 파일(HWP, DOCX, PDF 등) ➔ 노션 마크다운 자동 변환 (누적 및 버전 증가)**:
  ```bash
  python scripts/notion_ppt_bridge.py convert-script "<원고명 또는 Page ID>"
  ```
  → 첨부된 원본파일 다운로드 및 `원고변환지침` 적용, 노션 본문 하단에 새 버전 누적 Append 및 `버전` 속성 자동 증가
- **원고/지침 수집 및 이미지 로컬 자동 캐싱**:
  ```bash
  python scripts/notion_ppt_bridge.py fetch "<프로젝트명 또는 Page ID>"
  ```
  → 노션 본문 내 S3 이미지를 `md/<프로젝트명>/`로 다운로드하고 상대경로로 치환하여 `md/<프로젝트명>_v1.md` 및 `_guidelines.md` 생성
- **생성 결과 배포 및 노션 속성 업데이트**:
  ```bash
  python scripts/notion_ppt_bridge.py deploy "<Page ID>" "<pptx파일경로1>" "<pptx파일경로2>" --status "PPT 생성완료"
  ```
  → 구글 드라이브 `<프로젝트명>/` 하위 폴더로 복사 및 `PPT결과` 속성에 `[일시] 파일명` 누적, `진행상태`를 `PPT 생성완료`로 갱신

---

## 5. PPTX 제작 및 품질 관리 표준

### 5.1 ppt-master 스킬 사용 규칙 (Quick Generate 기본)
- 진입 시 `SKILL_DIR`은 `ppt-master/skills/ppt-master`로 고정한다.
- 프로젝트 초기화 시 항상 **Quick Generate 프로파일 및 4:3 비율(`--format ppt43`)**을 기본으로 사용한다:
  ```bash
  python ppt-master/skills/ppt-master/scripts/project_manager.py init <project_name> --format ppt43 --quick-generate
  ```
- 확정된 최종 md(`md/<프로젝트명>_v2.md`)와 캐싱된 이미지를 소스로 삼아 Quick Generate 절차(소스/리소스 준비 ➔ SVG 직접 저작 ➔ 품질 게이트 통과 ➔ export)를 따른다.

### 5.2 디자인 및 제작 규격 준수 (`spec/pptx-definition.md`)
- **강의 헤더 보존**: 상단 4줄 헤더(과목명, 차시, 슬라이드 번호, 슬라이드 제목) 영역을 반드시 정확히 분리·배치한다.
- **카드 및 인포그래픽 구조화**: 텍스트 단순 나열을 금지하고, 2단~4단 카드, 비교 대조 박스, 단계별 로드맵, 핵심 키워드 뱃지 등 인포그래픽 컴포넌트로 시각화한다.
- **스타일 및 톤앤매너**: `PPT-Guidelines`에서 선택된 스타일(스위스 미니멀, 다크 테크, 소프트 라운디드 등)의 컬러 팔레트와 라운딩 규칙을 엄격히 적용한다.

### 5.3 SVG 품질 검사 게이트 (`svg_quality_checker.py`) 필수 검증
모든 슬라이드 SVG 저작 후 반드시 품질 검사기를 통과(0 Errors)해야만 PPTX로 export할 수 있다:
```bash
python ppt-master/skills/ppt-master/scripts/svg_quality_checker.py <project_dir> --quick-generate --canonical-authoring --stage final --json
```
- **검증 항목**:
  1. `viewBox`: `0 0 960 720` (4:3 규격 일치 여부)
  2. `spec_lock.md` 동기화: 사용된 폰트 크기(px 단위 없는 순수 숫자), 컬러 팔레트가 선언과 일치하는지 확인
  3. 명도 대비(Contrast Ratio): 배경 대비 텍스트 최소 4.5:1 이상 충족
  4. 텍스트 박스 오버플로우 방지 및 안전 여백 확보

### 5.4 16:9 와이드 변환본 자동 생성
4:3 기본형 PPTX가 `exports/<project_name>_<timestamp>.pptx`로 export되면, 곧바로 `scripts/convert-ratio.py`를 실행하여 16:9 비율의 PPTX(`<project_name>_<timestamp>-169.pptx`)를 동일한 `exports/` 폴더에 생성한다:
```bash
python scripts/convert-ratio.py "<exports_경로>/<project_name>_<timestamp>.pptx"
```

---

## 6. 전체 실행 라이프사이클 요약

1. **원고 변환**: 사용자가 `PPT-Scripts`에 원본파일 등록 ➔ `convert-script` 실행 ➔ 마크다운 누적 Append 및 버전 갱신
2. **프로젝트 설정**: 사용자가 `PPT-Pipeline`에 새 프로젝트 등록 (원고, 개발지침 매핑, 상태를 `준비 완료`로 설정)
3. **AI 트리거**: 사용자가 대화창에 `"<프로젝트명> PPT 개발해줘"` 요청
4. **수집 및 캐싱**: AI가 `fetch` 실행 ➔ 마크다운 및 이미지 로컬 저장 (`md/`)
5. **정제 및 저작**: `spec/md-definition.md` 기반 정제 ➔ SVG 직접 저작 ➔ `svg_quality_checker` 게이트 통과 (0 Errors)
6. **빌드 및 16:9 변환**: 4:3 PPTX 빌드 ➔ `convert-ratio.py`로 16:9 변환본 동시 생성
7. **배포 및 피드백**: `deploy` 실행 ➔ Google Drive `<프로젝트명>/` 폴더로 복사 ➔ Notion `진행상태`를 `PPT 생성완료`로 갱신 및 `PPT결과`에 파일명 기록
8. **보고**: 사용자에게 로컬 산출물 경로 및 Google Drive 링크 안내
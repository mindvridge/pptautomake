# CLAUDE.md - PPT AutoMake 프로젝트 가이드

## 프로젝트 개요

PowerPoint 도식화 재구성 플러그인. PPTX 파일 내 이미지 도식(플로차트, 프로세스 등)을 분석하여 python-pptx 네이티브 요소(도형, 표, 차트, 커넥터)로 자동 변환한다.

## 기술 스택

- **언어**: Python 3.10+, HTML/CSS/JavaScript
- **핵심 라이브러리**: python-pptx, Pillow, anthropic (Claude Vision API), lxml, PyYAML
- **PPT 연동**: pywin32 (COM API), VBA 매크로
- **웹 서버 (선택)**: Flask, Flask-CORS, Office.js
- **테스트**: pytest

## 프로젝트 구조

```
pptautomake/
├── src/
│   ├── analyzer.py          # Module 1: SlideAnalyzer - PPTX 파싱, 슬라이드별 요소 추출
│   ├── classifier.py        # Module 2: ElementClassifier - 이미지/도형 분류 (도식 vs 사진)
│   ├── diagram_analyzer.py  # Module 3: DiagramAnalyzer - Claude Vision API로 도식 상세 분석
│   ├── builder.py           # Module 4: NativePPTXBuilder - 네이티브 PPTX 요소 생성 (핵심)
│   ├── composer.py          # Module 5: SlideComposer - 슬라이드 조합 및 출력
│   ├── main.py              # CLI 진입점
│   ├── ppt_plugin.py        # PowerPoint COM 플러그인 (pywin32, 서버 불필요)
│   ├── api.py               # Flask 웹 API 서버 (Web Add-in 백엔드, 선택)
│   └── utils/
│       ├── color_utils.py   # HEX/RGB 변환, 대비색 계산
│       ├── layout_utils.py  # 그리드, 원형, 선형 배치 계산
│       └── xml_utils.py     # PPTX XML 파싱 (OOXML 네임스페이스)
├── tests/
│   ├── test_analyzer.py
│   ├── test_classifier.py
│   ├── test_builder.py
│   └── fixtures/            # 테스트용 PPTX 샘플 (미구성)
├── addin/
│   ├── PPTAutoMake.bas      # VBA 매크로 모듈 (PPT 리본 버튼용)
│   ├── manifest.xml         # Office Web Add-in 매니페스트 (선택)
│   ├── taskpane.html        # Task Pane HTML (선택)
│   └── static/
│       ├── taskpane.css     # Task Pane 스타일
│       └── taskpane.js      # Task Pane 클라이언트 JS
├── docs/
├── config.yaml              # 분석/빌더/출력 설정
├── run_in_ppt.bat           # PPT 직접 실행 (더블클릭, 서버 불필요)
├── install_macro.bat        # VBA 매크로 설치 도우미
├── run_server.py            # Web Add-in 백엔드 서버 (선택)
├── requirements.txt
└── pptx-plugin-prompt.docx  # 기획 문서
```

## 파이프라인 흐름

```
[1] SlideAnalyzer → [2] ElementClassifier → [3] DiagramAnalyzer → [4] NativePPTXBuilder → [5] SlideComposer
    PPTX 파싱         도식/사진 분류        Vision API 분석       네이티브 변환          최종 조합
```

## 주요 실행 명령어

```bash
# 의존성 설치
pip install -r requirements.txt

# 전체 파이프라인 실행
python -m src.main input.pptx -o output.pptx --mode both

# 분석만 (변환 없이)
python -m src.main input.pptx --mode analyze

# 특정 슬라이드만 처리
python -m src.main input.pptx --slides 1,3,5-8

# 드라이런 (분석 결과만 출력)
python -m src.main input.pptx --dry-run

# 상세 로그
python -m src.main input.pptx -v

# PowerPoint 직접 연동 (서버 불필요, Windows)
run_in_ppt.bat                              # 전체 슬라이드 처리
python -m src.ppt_plugin --slides 7,8       # 특정 슬라이드만
python -m src.ppt_plugin --analyze          # 분석만
```

## PowerPoint 연동 (COM 방식, 권장)

서버 없이 PowerPoint에서 직접 실행:

1. PowerPoint에서 프레젠테이션을 연다
2. `run_in_ppt.bat` 더블클릭 (또는 `python -m src.ppt_plugin`)
3. 자동으로 현재 PPT를 분석 → 변환 → 슬라이드 삽입

### VBA 매크로 설치 (PPT 안에서 버튼으로 실행)

1. `install_macro.bat` 더블클릭
2. PowerPoint에서 Alt+F11 → 삽입 → 모듈
3. `addin/PPTAutoMake.bas` 내용 붙여넣기
4. Alt+F8 → `PPTAutoMake_Run` 실행

## Office Web Add-in (선택)

웹 기반 Add-in이 필요한 경우:

1. `python run_server.py`로 서버 시작
2. PowerPoint → 삽입 → 내 추가 기능 → 사용자 지정 추가 기능 업로드
3. `addin/manifest.xml` 선택
4. 리본 메뉴 홈 탭에 "도식 변환" 버튼 표시됨

## 테스트

```bash
# 전체 테스트
pytest tests/ -v

# 개별 모듈 테스트
pytest tests/test_analyzer.py -v
pytest tests/test_classifier.py -v
pytest tests/test_builder.py -v
```

## 주요 데이터 클래스

- `SlideElement` (analyzer.py) - 슬라이드 내 개별 요소 메타데이터
- `SlideAnalysisResult` (analyzer.py) - 슬라이드 분석 결과
- `ClassifiedElement` (classifier.py) - 분류된 요소 (ContentType, DiagramType 포함)
- `DiagramData` (diagram_analyzer.py) - 도식 상세 분석 결과 (노드, 연결, 스타일)
- `DiagramType` (classifier.py) - 10종: flowchart, process, comparison_table, hierarchy, cycle, grid, timeline, bar_chart, pie_chart, infographic
- `ContentType` (classifier.py) - 8종: diagram_image, photo, group_diagram, table, chart, text_only, connector, unknown

## 설정 (config.yaml)

- `analysis.vision_model` - Claude Vision API 모델명
- `analysis.min_diagram_area_ratio` - 도식 판별 최소 면적 비율 (기본 0.25)
- `builder.default_colors` - 도식 기본 색상 (primary, secondary, accent)
- `builder.default_fonts` - 폰트 설정 (맑은 고딕)
- `output.insert_mode` - 슬라이드 삽입 위치 (after_original / append_end / replace)

## 코딩 컨벤션

- 모든 모듈/함수 docstring은 한국어로 작성
- 타입 힌트 사용 (`from __future__ import annotations`)
- 모듈 간 import는 `src.` 접두사 사용
- 로깅: `logging.getLogger(__name__)` 패턴
- 에러 시 기본값 반환 (DiagramData(), {}, []) - 파이프라인 중단 방지
- PPTX XML 처리 시 `src/utils/xml_utils.py`의 네임스페이스 상수 사용

## 외부 API 의존성

- **Claude Vision API** (anthropic 패키지): DiagramAnalyzer에서 이미지 도식 분석에 사용
- API 키는 환경변수 `ANTHROPIC_API_KEY`로 설정

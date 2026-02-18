"""PowerPoint COM 플러그인 - 서버 없이 PPT에서 직접 실행

pywin32의 win32com을 사용하여 현재 열린 PowerPoint 프레젠테이션에
직접 접근하고, 도식 분석/변환 파이프라인을 실행한 후
결과 슬라이드를 현재 프레젠테이션에 자동 삽입한다.

Usage:
    python -m src.ppt_plugin              # 전체 슬라이드 처리
    python -m src.ppt_plugin --slides 7,8 # 특정 슬라이드만
    python -m src.ppt_plugin --analyze    # 분석만 (변환 없이)
"""

from __future__ import annotations

import argparse
import logging
import shutil
import sys
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


def _check_win32com():
    """pywin32 설치 여부를 확인한다."""
    try:
        import win32com.client
        return True
    except ImportError:
        print('[오류] pywin32가 설치되어 있지 않습니다.')
        print('       pip install pywin32')
        return False


def _check_ollama_ready(analysis_config: dict) -> bool:
    """Ollama 서버 상태 확인 및 비전 모델 자동 다운로드를 수행한다."""
    model_name = analysis_config.get('ollama_model', 'llama3.2-vision')
    base_url = analysis_config.get('ollama_base_url', 'http://localhost:11434')

    try:
        import ollama
    except ImportError:
        print('[오류] ollama 패키지가 설치되지 않았습니다.')
        print('       pip install ollama')
        return False

    # 1) 서버 연결 확인
    try:
        client = ollama.Client(host=base_url)
        models = client.list()
    except Exception:
        print(f'[오류] Ollama 서버에 연결할 수 없습니다 ({base_url})')
        print('       Ollama를 먼저 실행해주세요: https://ollama.com')
        return False

    # 2) 모델 존재 여부 확인
    model_names = [m.model for m in models.models] if models.models else []
    # "llama3.2-vision:latest" 등 태그 포함 비교
    found = any(
        m == model_name or m.startswith(f'{model_name}:')
        for m in model_names
    )

    if found:
        print(f'  -> 비전 모델 확인: {model_name}')
        return True

    # 3) 모델 자동 다운로드
    print(f'  -> 비전 모델 ({model_name})이 없습니다. 자동 다운로드 중...')
    print(f'     (최초 1회만 필요, 수 분 소요될 수 있습니다)')
    try:
        client.pull(model_name)
        print(f'  -> 다운로드 완료: {model_name}')
        return True
    except Exception as e:
        print(f'[오류] 모델 다운로드 실패: {e}')
        print(f'       수동으로 실행: ollama pull {model_name}')
        return False


def get_powerpoint_app():
    """실행 중인 PowerPoint 인스턴스에 연결한다."""
    import win32com.client

    try:
        ppt = win32com.client.GetActiveObject('PowerPoint.Application')
        return ppt
    except Exception:
        print('[오류] 실행 중인 PowerPoint를 찾을 수 없습니다.')
        print('       PowerPoint를 먼저 열고 프레젠테이션을 열어주세요.')
        return None


def get_active_presentation(ppt_app):
    """현재 활성 프레젠테이션을 가져온다."""
    try:
        if ppt_app.Presentations.Count == 0:
            print('[오류] 열린 프레젠테이션이 없습니다.')
            return None
        return ppt_app.ActivePresentation
    except Exception as e:
        print(f'[오류] 프레젠테이션에 접근할 수 없습니다: {e}')
        return None


def save_presentation_to_temp(presentation) -> str:
    """현재 프레젠테이션을 임시 파일로 저장한다."""
    temp_dir = Path(tempfile.gettempdir()) / 'pptautomake'
    temp_dir.mkdir(exist_ok=True)
    temp_path = str(temp_dir / 'current_presentation.pptx')
    presentation.SaveCopyAs(temp_path)

    if not Path(temp_path).exists():
        raise FileNotFoundError(f'프레젠테이션 저장 실패: {temp_path}')
    return temp_path


def create_backup(presentation) -> str | None:
    """현재 프레젠테이션의 백업을 생성한다."""
    try:
        temp_dir = Path(tempfile.gettempdir()) / 'pptautomake' / 'backups'
        temp_dir.mkdir(parents=True, exist_ok=True)
        name = Path(presentation.Name).stem
        backup_path = str(temp_dir / f'{name}_backup.pptx')
        presentation.SaveCopyAs(backup_path)
        return backup_path
    except Exception as e:
        logger.warning('백업 생성 실패: %s', e)
        return None


def insert_slides_from_file(presentation, rebuilt_path: str) -> bool:
    """변환된 PPTX의 슬라이드를 현재 프레젠테이션 끝에 삽입한다."""
    try:
        last_slide = presentation.Slides.Count
        presentation.Slides.InsertFromFile(rebuilt_path, last_slide)
        new_count = presentation.Slides.Count
        inserted = new_count - last_slide
        print(f'  -> {inserted}개 슬라이드가 프레젠테이션 끝에 삽입되었습니다.')
        return True
    except Exception as e:
        logger.warning('슬라이드 삽입 실패: %s', e)
        print(f'[경고] 슬라이드 자동 삽입에 실패했습니다: {e}')
        print(f'       변환된 파일을 직접 열어주세요: {rebuilt_path}')
        return False


def run_pipeline(
    temp_path: str,
    config: dict,
    slide_indices: list[int] | None = None,
    analyze_only: bool = False,
) -> tuple[str | None, dict]:
    """파이프라인을 실행한다.

    Returns:
        (output_path, summary) - 변환된 파일 경로와 결과 요약
    """
    from src.analyzer import SlideAnalyzer
    from src.classifier import ElementClassifier
    from src.diagram_analyzer import DiagramAnalyzer
    from src.builder import NativePPTXBuilder
    from src.composer import SlideComposer

    summary = {'total_slides': 0, 'rebuild_count': 0, 'details': []}

    # Stage 1: 슬라이드 분석
    print('[1/4] 슬라이드 분석 중...')
    analyzer = SlideAnalyzer()
    analyses = analyzer.analyze(temp_path)

    if slide_indices:
        analyses = [a for a in analyses if a.slide_index in slide_indices]

    summary['total_slides'] = len(analyses)
    print(f'  -> {len(analyses)}개 슬라이드 분석 완료')

    # Stage 2: 요소 분류
    print('[2/4] 요소 분류 중...')
    classifier = ElementClassifier(config.get('analysis', {}))
    classifications = classifier.classify_all(analyses)

    rebuild_count = sum(1 for c in classifications if c.needs_rebuild)
    summary['rebuild_count'] = rebuild_count
    print(f'  -> 재구성 필요 슬라이드: {rebuild_count}개')

    if analyze_only:
        _print_analysis(analyses, classifications)
        return None, summary

    if rebuild_count == 0:
        print('\n재구성할 도식이 없습니다.')
        return None, summary

    # Stage 3: 도식 상세 분석
    analysis_config = config.get('analysis', {})
    backend = analysis_config.get('vision_backend', 'local')
    if backend == 'local':
        model_name = analysis_config.get('local_model', 'Qwen/Qwen2.5-VL-7B-Instruct')
        print(f'[3/4] 도식 상세 분석 중 (로컬 모델: {model_name})...')
        print(f'      (최초 실행 시 모델 자동 다운로드, 이후 캐시 사용)')
    elif backend == 'ollama':
        model_name = analysis_config.get('ollama_model', 'llama3.2-vision')
        print(f'[3/4] 도식 상세 분석 중 (Ollama: {model_name})...')
        _check_ollama_ready(analysis_config)
    else:
        print('[3/4] 도식 상세 분석 중 (Anthropic API)...')
    diagram_analyzer = DiagramAnalyzer(analysis_config)
    diagram_results = {}

    total_diagrams = sum(
        len(c.diagram_elements) for c in classifications if c.needs_rebuild
    )
    processed = 0

    for classification in classifications:
        if not classification.needs_rebuild:
            continue
        slide_diagrams = []
        for ce in classification.diagram_elements:
            processed += 1
            print(f'  [{processed}/{total_diagrams}] 슬라이드 {classification.slide_index + 1} 분석 중...')
            diagram_data = diagram_analyzer.analyze(ce)
            slide_diagrams.append((ce, diagram_data))
            detail = {
                'slide': classification.slide_index + 1,
                'diagram_type': diagram_data.diagram_type,
                'node_count': diagram_data.node_count,
            }
            summary['details'].append(detail)
            print(f'    -> {detail["diagram_type"]} ({detail["node_count"]}개 노드)')
        if slide_diagrams:
            diagram_results[classification.slide_index] = slide_diagrams

    if not diagram_results:
        print('\n도식 분석 결과가 없습니다.')
        return None, summary

    # Stage 4: 네이티브 빌드 + 조합
    print('[4/4] 네이티브 요소 재구성 중...')
    builder = NativePPTXBuilder(config.get('builder', {}))
    composer = SlideComposer(config)

    temp_dir = Path(tempfile.gettempdir()) / 'pptautomake'
    output_path = str(temp_dir / 'rebuilt_output.pptx')

    composer.compose(
        temp_path, analyses, classifications,
        diagram_results, output_path, builder,
    )

    print(f'  -> 변환 완료: {output_path}')
    return output_path, summary


def _print_analysis(analyses, classifications):
    """분석 결과를 출력한다."""
    print('\n' + '=' * 50)
    print('  분석 결과')
    print('=' * 50)

    for analysis, classification in zip(analyses, classifications):
        idx = analysis.slide_index + 1
        total = len(analysis.elements)
        rebuild = 'O' if classification.needs_rebuild else 'X'

        print(f'\n  슬라이드 {idx} (요소: {total}개, 재구성: {rebuild})')

        type_counts = {}
        for elem in analysis.elements:
            t = elem.element_type
            type_counts[t] = type_counts.get(t, 0) + 1
        for t, count in type_counts.items():
            print(f'    - {t}: {count}개')

        if classification.needs_rebuild:
            for ce in classification.diagram_elements:
                print(f'    * 도식: {ce.content_type.value} (신뢰도: {ce.confidence:.0%})')

    print('\n' + '=' * 50)


def main():
    parser = argparse.ArgumentParser(
        description='PowerPoint 도식 자동변환 - PPT 직접 연동 (서버 불필요)',
    )
    parser.add_argument(
        '--slides', default=None,
        help='처리할 슬라이드 번호 (예: 1,3,5-8)',
    )
    parser.add_argument(
        '--analyze', action='store_true',
        help='분석만 수행 (변환 없이)',
    )
    parser.add_argument(
        '--no-insert', action='store_true',
        help='변환 후 슬라이드 자동 삽입 안 함',
    )
    parser.add_argument(
        '-v', '--verbose', action='store_true',
        help='상세 로그 출력',
    )
    args = parser.parse_args()

    # 로깅 설정
    log_level = logging.DEBUG if args.verbose else logging.WARNING
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    )

    print()
    print('=' * 50)
    print('  PPT 도식 자동변환')
    print('  이미지 도식 → 네이티브 요소')
    print('=' * 50)
    print()

    # pywin32 확인
    if not _check_win32com():
        sys.exit(1)

    # PowerPoint 연결
    print('[준비] PowerPoint 연결 중...')
    ppt_app = get_powerpoint_app()
    if not ppt_app:
        sys.exit(1)

    presentation = get_active_presentation(ppt_app)
    if not presentation:
        sys.exit(1)

    ppt_name = presentation.Name
    print(f'  -> 연결됨: {ppt_name}')
    print(f'  -> 슬라이드 수: {presentation.Slides.Count}')

    # 백업 생성
    total_slides = presentation.Slides.Count
    print('[준비] 백업 생성 중...')
    backup_path = create_backup(presentation)
    if backup_path:
        print(f'  -> 백업: {backup_path}')
    else:
        print('  -> 백업 생성 실패 (계속 진행)')

    # 임시 파일로 저장
    print('[준비] 프레젠테이션 복사 중...')
    temp_path = save_presentation_to_temp(presentation)
    print(f'  -> {temp_path}')

    # 슬라이드 범위 파싱
    slide_indices = None
    if args.slides:
        from src.main import parse_slide_ranges
        slide_indices = parse_slide_ranges(args.slides, max_slide=total_slides)
        if not slide_indices:
            print(f'[오류] 유효한 슬라이드 번호가 없습니다 (전체: {total_slides}개)')
            sys.exit(1)
        print(f'  -> 처리 대상: 슬라이드 {[i + 1 for i in slide_indices]}')

    print()

    # 설정 로드
    from src.main import load_config
    config = load_config()

    # 파이프라인 실행
    output_path, summary = run_pipeline(
        temp_path, config, slide_indices, args.analyze,
    )

    if output_path and not args.no_insert:
        # 변환된 슬라이드를 현재 PPT에 삽입
        print('\n[삽입] 변환된 슬라이드를 현재 프레젠테이션에 삽입 중...')
        insert_slides_from_file(presentation, output_path)

    print('\n완료!')


if __name__ == '__main__':
    main()

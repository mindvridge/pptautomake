"""메인 실행 파일 - CLI 인터페이스

PowerPoint 도식화 재구성 플러그인의 진입점.
PPTX 파일을 분석하고 이미지 도식을 네이티브 요소로 자동 변환한다.

Usage:
    python src/main.py input.pptx -o output.pptx --mode both
    python src/main.py input.pptx --dry-run
    python src/main.py input.pptx --slides 7,8,10
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

import yaml

from src.analyzer import SlideAnalyzer
from src.classifier import ElementClassifier, ContentType
from src.diagram_analyzer import DiagramAnalyzer
from src.builder import NativePPTXBuilder
from src.composer import SlideComposer

logger = logging.getLogger(__name__)


def load_config(config_path: str | None = None) -> dict:
    """설정 파일을 로드한다."""
    if config_path is None:
        # 프로젝트 루트의 config.yaml 탐색
        candidates = [
            Path(__file__).parent.parent / 'config.yaml',
            Path.cwd() / 'config.yaml',
        ]
        for candidate in candidates:
            if candidate.exists():
                config_path = str(candidate)
                break

    if config_path and Path(config_path).exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    return {}


def parse_slide_ranges(slides_str: str, max_slide: int | None = None) -> list[int]:
    """슬라이드 번호 문자열을 파싱한다.

    예: "1,3,5-8" -> [0, 2, 4, 5, 6, 7] (0-indexed)

    Args:
        slides_str: 슬라이드 번호 문자열 (1-indexed)
        max_slide: 최대 슬라이드 번호 (1-indexed). 초과 시 경고 후 제외.
    """
    indices = []
    skipped = []
    for part in slides_str.split(','):
        part = part.strip()
        if not part:
            continue
        try:
            if '-' in part:
                start, end = part.split('-', 1)
                for i in range(int(start), int(end) + 1):
                    if i < 1:
                        continue
                    if max_slide and i > max_slide:
                        skipped.append(i)
                        continue
                    indices.append(i - 1)
            else:
                val = int(part)
                if val < 1:
                    continue
                if max_slide and val > max_slide:
                    skipped.append(val)
                    continue
                indices.append(val - 1)
        except ValueError:
            logger.warning("잘못된 슬라이드 번호 무시: '%s'", part)

    if skipped:
        logger.warning(
            "존재하지 않는 슬라이드 번호 무시: %s (최대: %d)",
            skipped, max_slide,
        )
    return sorted(set(indices))


def run_analysis(
    pptx_path: str,
    config: dict,
    slide_indices: list[int] | None = None,
) -> tuple:
    """[1/4] 슬라이드 분석 + [2/4] 요소 분류를 수행한다."""
    print("[1/4] 슬라이드 분석 중...")
    analyzer = SlideAnalyzer()
    analyses = analyzer.analyze(pptx_path)

    # 특정 슬라이드만 처리
    if slide_indices:
        analyses = [
            a for a in analyses if a.slide_index in slide_indices
        ]

    print(f"  -> {len(analyses)}개 슬라이드 분석 완료")

    print("[2/4] 요소 분류 중...")
    classifier = ElementClassifier(config.get('analysis', {}))
    classifications = classifier.classify_all(analyses)

    rebuild_count = sum(1 for c in classifications if c.needs_rebuild)
    print(f"  -> 재구성 필요 슬라이드: {rebuild_count}개")

    return analyses, classifications


def run_diagram_analysis(
    classifications: list,
    config: dict,
) -> dict:
    """[3/4] 도식 상세 분석을 수행한다."""
    print("[3/4] 도식 상세 분석 중...")
    diagram_analyzer = DiagramAnalyzer(config.get('analysis', {}))

    diagram_results = {}  # {slide_index: [(classified_element, diagram_data), ...]}

    for classification in classifications:
        if not classification.needs_rebuild:
            continue

        slide_diagrams = []
        for ce in classification.diagram_elements:
            diagram_data = diagram_analyzer.analyze(ce)
            slide_diagrams.append((ce, diagram_data))
            print(
                f"  -> 슬라이드 {classification.slide_index + 1}: "
                f"{diagram_data.diagram_type} ({diagram_data.node_count}개 노드)"
            )

        if slide_diagrams:
            diagram_results[classification.slide_index] = slide_diagrams

    print(f"  -> {len(diagram_results)}개 슬라이드 도식 분석 완료")
    return diagram_results


def run_rebuild(
    pptx_path: str,
    output_path: str | None,
    analyses: list,
    classifications: list,
    diagram_results: dict,
    config: dict,
) -> str:
    """[4/4] 네이티브 요소 재구성을 수행한다."""
    print("[4/4] 네이티브 요소 재구성 중...")
    builder = NativePPTXBuilder(config.get('builder', {}))
    composer = SlideComposer(config)

    output = composer.compose(
        pptx_path, analyses, classifications,
        diagram_results, output_path, builder,
    )
    print(f"  -> 출력 파일: {output}")
    return output


def print_analysis_report(
    analyses: list,
    classifications: list,
    diagram_results: dict | None = None,
) -> None:
    """분석 결과 보고서를 출력한다."""
    print("\n" + "=" * 60)
    print("분석 결과 보고서")
    print("=" * 60)

    for analysis, classification in zip(analyses, classifications):
        idx = analysis.slide_index + 1
        total = len(analysis.elements)
        rebuild = "O" if classification.needs_rebuild else "X"

        print(f"\n슬라이드 {idx} (요소: {total}개, 재구성: {rebuild})")

        type_counts = {}
        for elem in analysis.elements:
            t = elem.element_type
            type_counts[t] = type_counts.get(t, 0) + 1

        for t, count in type_counts.items():
            print(f"  - {t}: {count}개")

        if classification.needs_rebuild:
            for ce in classification.diagram_elements:
                print(f"  * 도식 발견: {ce.content_type.value} (신뢰도: {ce.confidence:.1%})")

        if diagram_results and analysis.slide_index in diagram_results:
            for ce, dd in diagram_results[analysis.slide_index]:
                print(f"  * 분석 결과: {dd.diagram_type} ({dd.node_count}개 노드)")

    print("\n" + "=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description='PowerPoint 도식화 재구성 플러그인 - PPTX 내 이미지 도식을 네이티브 요소로 자동 변환',
    )
    parser.add_argument(
        'input',
        help='입력 PPTX 파일 경로',
    )
    parser.add_argument(
        '-o', '--output',
        default=None,
        help='출력 PPTX 파일 경로 (기본: input_rebuilt.pptx)',
    )
    parser.add_argument(
        '--slides',
        default=None,
        help='처리할 슬라이드 번호 (예: 1,3,5-8)',
    )
    parser.add_argument(
        '--mode',
        choices=['analyze', 'rebuild', 'both'],
        default='both',
        help='실행 모드 (기본: both)',
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='실제 변환 없이 분석 결과만 출력',
    )
    parser.add_argument(
        '--config',
        default=None,
        help='설정 파일 경로 (기본: config.yaml)',
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )

    args = parser.parse_args()

    # 로깅 설정
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    )

    # 입력 파일 확인
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"오류: 입력 파일을 찾을 수 없습니다: {args.input}", file=sys.stderr)
        sys.exit(1)
    if not input_path.suffix.lower() == '.pptx':
        print(f"오류: PPTX 파일만 지원합니다: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 설정 로드
    config = load_config(args.config)
    print(f"PowerPoint 도식화 재구성 플러그인")
    print(f"입력: {args.input}")

    # 슬라이드 범위 파싱
    slide_indices = None
    if args.slides:
        slide_indices = parse_slide_ranges(args.slides)
        print(f"처리 대상 슬라이드: {[i + 1 for i in slide_indices]}")

    # 실행
    if args.mode in ('analyze', 'both') or args.dry_run:
        analyses, classifications = run_analysis(
            args.input, config, slide_indices,
        )

        diagram_results = None
        if not args.dry_run or args.mode == 'analyze':
            diagram_results = run_diagram_analysis(classifications, config)

        print_analysis_report(analyses, classifications, diagram_results)

        if args.dry_run:
            print("\n--dry-run 모드: 변환을 수행하지 않습니다.")
            return

        if args.mode == 'both':
            if diagram_results:
                run_rebuild(
                    args.input, args.output,
                    analyses, classifications,
                    diagram_results, config,
                )
            else:
                print("재구성할 도식이 없습니다.")

    elif args.mode == 'rebuild':
        analyses, classifications = run_analysis(
            args.input, config, slide_indices,
        )
        diagram_results = run_diagram_analysis(classifications, config)
        if diagram_results:
            run_rebuild(
                args.input, args.output,
                analyses, classifications,
                diagram_results, config,
            )
        else:
            print("재구성할 도식이 없습니다.")

    print("\n완료!")


if __name__ == '__main__':
    main()

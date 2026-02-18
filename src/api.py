"""Flask 웹 API 서버

기존 CLI 파이프라인을 REST API로 래핑하여
Office Add-in Task Pane에서 호출할 수 있도록 한다.

Endpoints:
    POST /api/analyze   - PPTX 분석 (분류까지)
    POST /api/process   - 전체 파이프라인 (분석 + 변환)
    GET  /api/download/<filename> - 결과 파일 다운로드
"""

from __future__ import annotations

import base64
import logging
import os
import tempfile
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import yaml

from src.analyzer import SlideAnalyzer
from src.classifier import ElementClassifier
from src.diagram_analyzer import DiagramAnalyzer
from src.builder import NativePPTXBuilder
from src.composer import SlideComposer

logger = logging.getLogger(__name__)

app = Flask(
    __name__,
    static_folder=str(Path(__file__).parent.parent / 'addin' / 'static'),
    static_url_path='/static',
)
CORS(app)

# 임시 파일 저장 디렉토리
UPLOAD_DIR = Path(tempfile.gettempdir()) / 'pptautomake_uploads'
OUTPUT_DIR = Path(tempfile.gettempdir()) / 'pptautomake_outputs'
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


_CLEANUP_MAX_AGE = 3600  # 1시간 이상 된 임시 파일 삭제


def _cleanup_old_files():
    """1시간 이상 지난 임시 파일을 삭제한다."""
    now = time.time()
    for directory in (UPLOAD_DIR, OUTPUT_DIR):
        try:
            for f in directory.iterdir():
                if f.is_file() and (now - f.stat().st_mtime) > _CLEANUP_MAX_AGE:
                    f.unlink(missing_ok=True)
                    logger.debug("임시 파일 삭제: %s", f)
        except Exception as e:
            logger.debug("임시 파일 정리 실패: %s", e)


def _load_config() -> dict:
    """config.yaml 로드"""
    config_path = Path(__file__).parent.parent / 'config.yaml'
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    return {}


@app.route('/')
def index():
    """Task Pane HTML 제공"""
    taskpane_path = Path(__file__).parent.parent / 'addin' / 'taskpane.html'
    return send_file(taskpane_path)


@app.route('/api/health', methods=['GET'])
def health():
    """서버 상태 확인"""
    return jsonify({'status': 'ok'})


@app.route('/api/analyze', methods=['POST'])
def analyze():
    """PPTX 파일을 분석하여 슬라이드별 도식 정보를 반환한다."""
    _cleanup_old_files()

    if 'file' not in request.files:
        return jsonify({'error': '파일이 첨부되지 않았습니다.'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.pptx'):
        return jsonify({'error': 'PPTX 파일만 지원합니다.'}), 400

    slides_str = request.form.get('slides', '')

    # 파일 저장
    file_id = str(uuid.uuid4())[:8]
    input_path = UPLOAD_DIR / f'{file_id}_{file.filename}'
    file.save(str(input_path))

    try:
        config = _load_config()

        # Stage 1: 슬라이드 분석
        analyzer = SlideAnalyzer()
        analyses = analyzer.analyze(str(input_path))

        # 슬라이드 필터링
        if slides_str:
            from src.main import parse_slide_ranges
            slide_indices = parse_slide_ranges(slides_str)
            analyses = [a for a in analyses if a.slide_index in slide_indices]

        # Stage 2: 요소 분류
        classifier = ElementClassifier(config.get('analysis', {}))
        classifications = classifier.classify_all(analyses)

        # 결과 구성
        result = {
            'file_id': file_id,
            'filename': file.filename,
            'total_slides': len(analyses),
            'slides': [],
        }

        for analysis, classification in zip(analyses, classifications):
            slide_info = {
                'index': analysis.slide_index + 1,
                'element_count': len(analysis.elements),
                'needs_rebuild': classification.needs_rebuild,
                'elements': {},
                'diagrams': [],
            }

            # 요소 타입별 카운트
            for elem in analysis.elements:
                t = elem.element_type
                slide_info['elements'][t] = slide_info['elements'].get(t, 0) + 1

            # 도식 정보
            for ce in classification.diagram_elements:
                slide_info['diagrams'].append({
                    'content_type': ce.content_type.value,
                    'diagram_type': ce.diagram_type.value,
                    'confidence': round(ce.confidence, 2),
                })

            result['slides'].append(slide_info)

        rebuild_count = sum(1 for s in result['slides'] if s['needs_rebuild'])
        result['rebuild_count'] = rebuild_count

        return jsonify(result)

    except Exception as e:
        logger.exception('분석 중 오류 발생')
        return jsonify({'error': f'분석 중 오류가 발생했습니다: {str(e)}'}), 500


@app.route('/api/process', methods=['POST'])
def process():
    """전체 파이프라인을 실행한다 (분석 + 변환)."""
    _cleanup_old_files()

    if 'file' not in request.files:
        return jsonify({'error': '파일이 첨부되지 않았습니다.'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.pptx'):
        return jsonify({'error': 'PPTX 파일만 지원합니다.'}), 400

    slides_str = request.form.get('slides', '')
    return_base64 = request.form.get('return_base64', '') == 'true'

    # 파일 저장
    file_id = str(uuid.uuid4())[:8]
    input_path = UPLOAD_DIR / f'{file_id}_{file.filename}'
    file.save(str(input_path))

    try:
        config = _load_config()

        # Stage 1: 슬라이드 분석
        analyzer = SlideAnalyzer()
        analyses = analyzer.analyze(str(input_path))

        if slides_str:
            from src.main import parse_slide_ranges
            slide_indices = parse_slide_ranges(slides_str)
            analyses = [a for a in analyses if a.slide_index in slide_indices]

        # Stage 2: 요소 분류
        classifier = ElementClassifier(config.get('analysis', {}))
        classifications = classifier.classify_all(analyses)

        # Stage 3: 도식 상세 분석
        diagram_analyzer = DiagramAnalyzer(config.get('analysis', {}))
        diagram_results = {}

        for classification in classifications:
            if not classification.needs_rebuild:
                continue
            slide_diagrams = []
            for ce in classification.diagram_elements:
                diagram_data = diagram_analyzer.analyze(ce)
                slide_diagrams.append((ce, diagram_data))
            if slide_diagrams:
                diagram_results[classification.slide_index] = slide_diagrams

        if not diagram_results:
            return jsonify({
                'file_id': file_id,
                'message': '재구성할 도식이 없습니다.',
                'download_url': None,
            })

        # Stage 4-5: 네이티브 빌드 + 조합
        builder = NativePPTXBuilder(config.get('builder', {}))
        composer = SlideComposer(config)

        output_filename = f'{file_id}_rebuilt.pptx'
        output_path = OUTPUT_DIR / output_filename

        composer.compose(
            str(input_path), analyses, classifications,
            diagram_results, str(output_path), builder,
        )

        # 결과 요약
        result = {
            'file_id': file_id,
            'message': '변환 완료',
            'download_url': f'/api/download/{output_filename}',
            'rebuilt_slides': len(diagram_results),
            'details': [],
        }

        for slide_idx, diagrams in diagram_results.items():
            for ce, dd in diagrams:
                result['details'].append({
                    'slide': slide_idx + 1,
                    'diagram_type': dd.diagram_type,
                    'node_count': dd.node_count,
                })

        # Office.js insertSlidesFromBase64 용 base64 응답
        if return_base64:
            with open(str(output_path), 'rb') as f:
                result['base64'] = base64.b64encode(f.read()).decode('ascii')

        return jsonify(result)

    except Exception as e:
        logger.exception('처리 중 오류 발생')
        return jsonify({'error': f'처리 중 오류가 발생했습니다: {str(e)}'}), 500


@app.route('/api/download/<filename>', methods=['GET'])
def download(filename: str):
    """변환된 PPTX 파일을 다운로드한다."""
    # Path traversal 방지: 파일명에서 디렉토리 구분자 제거
    safe_filename = Path(filename).name
    if not safe_filename or safe_filename != filename:
        return jsonify({'error': '잘못된 파일명입니다.'}), 400

    file_path = OUTPUT_DIR / safe_filename
    # OUTPUT_DIR 바깥 경로 접근 차단
    if not file_path.resolve().is_relative_to(OUTPUT_DIR.resolve()):
        return jsonify({'error': '잘못된 파일 경로입니다.'}), 400

    if not file_path.exists():
        return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404

    return send_file(
        str(file_path),
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=safe_filename,
    )


def create_app() -> Flask:
    """앱 팩토리"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    )
    return app


if __name__ == '__main__':
    create_app()
    port = int(os.environ.get('PORT', 5000))
    print(f'PPT AutoMake API 서버: http://localhost:{port}')
    print('HTTPS가 필요하면 run_server.py를 사용하세요.')
    app.run(host='0.0.0.0', port=port, debug=True)

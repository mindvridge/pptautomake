"""Module 5: SlideComposer - 슬라이드 조합 및 출력

분석된 요소를 조합하여 최종 슬라이드를 구성한다.
원본 슬라이드 뒤에 재구성된 슬라이드를 삽입한다.
"""

from __future__ import annotations

import logging
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from src.analyzer import SlideAnalysisResult
from src.classifier import (
    SlideClassificationResult,
    ClassifiedElement,
    ContentType,
)
from src.diagram_analyzer import DiagramData
from src.builder import NativePPTXBuilder

logger = logging.getLogger(__name__)


class SlideComposer:
    """재구성된 요소를 조합하여 최종 PPTX를 생성하는 모듈"""

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        output_config = self.config.get('output', {})
        self.insert_mode = output_config.get('insert_mode', 'after_original')
        self.keep_original = output_config.get('keep_original', True)
        self.output_suffix = output_config.get('output_suffix', '_rebuilt')

    def compose(
        self,
        pptx_path: str,
        analyses: list[SlideAnalysisResult],
        classifications: list[SlideClassificationResult],
        diagram_results: dict[int, list[tuple[ClassifiedElement, DiagramData]]],
        output_path: str | None = None,
        builder: NativePPTXBuilder | None = None,
    ) -> str:
        """원본 PPTX에 재구성된 슬라이드를 삽입하여 출력한다.

        Args:
            pptx_path: 원본 PPTX 파일 경로
            analyses: 슬라이드 분석 결과 리스트
            classifications: 슬라이드 분류 결과 리스트
            diagram_results: {slide_index: [(classified_element, diagram_data), ...]}
            output_path: 출력 파일 경로 (None이면 자동 생성)
            builder: NativePPTXBuilder 인스턴스

        Returns:
            출력 파일 경로
        """
        if output_path is None:
            p = Path(pptx_path)
            output_path = str(p.parent / f"{p.stem}{self.output_suffix}{p.suffix}")

        if builder is None:
            builder = NativePPTXBuilder(self.config.get('builder', {}))

        logger.info("슬라이드 조합 시작: %s -> %s", pptx_path, output_path)

        # 원본 PPTX 로드
        prs = Presentation(pptx_path)

        # 재구성이 필요한 슬라이드 식별
        rebuild_indices = set()
        for classification in classifications:
            if classification.needs_rebuild:
                rebuild_indices.add(classification.slide_index)

        logger.info(
            "재구성 대상 슬라이드: %s",
            sorted(i + 1 for i in rebuild_indices),
        )

        if not rebuild_indices:
            logger.info("재구성이 필요한 슬라이드가 없습니다")
            prs.save(output_path)
            return output_path

        # 슬라이드 삽입 (뒤에서부터 처리하여 인덱스 변경 방지)
        slides_to_insert = []
        for slide_idx in sorted(rebuild_indices, reverse=True):
            if slide_idx not in diagram_results:
                continue

            diagram_pairs = diagram_results[slide_idx]
            if not diagram_pairs:
                continue

            slides_to_insert.append((slide_idx, diagram_pairs))

        # 역순으로 처리하여 인덱스 유지
        for slide_idx, diagram_pairs in slides_to_insert:
            self._insert_rebuilt_slide(
                prs, slide_idx, diagram_pairs, builder, analyses
            )

        try:
            prs.save(output_path)
        except Exception as e:
            logger.error("출력 파일 저장 실패: %s - %s", output_path, e)
            raise
        logger.info("출력 파일 저장 완료: %s", output_path)
        return output_path

    def _insert_rebuilt_slide(
        self,
        prs: Presentation,
        slide_idx: int,
        diagram_pairs: list[tuple[ClassifiedElement, DiagramData]],
        builder: NativePPTXBuilder,
        analyses: list[SlideAnalysisResult],
    ) -> None:
        """재구성된 슬라이드를 원본 뒤에 삽입한다."""
        # 새 빈 슬라이드 추가
        slide_layout = prs.slide_layouts[6]  # 빈 레이아웃
        new_slide = prs.slides.add_slide(slide_layout)

        # 원본 슬라이드의 텍스트 요소 복사 (제목, 부제 등)
        if slide_idx < len(analyses):
            analysis = analyses[slide_idx]
            self._copy_text_elements(prs, slide_idx, new_slide, analysis)

        # 각 도식 데이터를 새 슬라이드에 빌드
        shapes_before = len(new_slide.shapes)
        built_count = 0
        for classified_elem, diagram_data in diagram_pairs:
            if diagram_data.nodes or diagram_data.table_data or diagram_data.chart_data:
                builder.build(new_slide, diagram_data)
                built_count += 1
            elif diagram_data.diagram_type != 'unknown':
                logger.warning(
                    "슬라이드 %d: 도식 '%s'에 빌드할 데이터가 없습니다",
                    slide_idx + 1, diagram_data.diagram_type,
                )

        shapes_after = len(new_slide.shapes)
        if built_count > 0 and shapes_after == shapes_before:
            logger.warning(
                "슬라이드 %d: 빌드를 실행했지만 도형이 생성되지 않았습니다",
                slide_idx + 1,
            )

        # 슬라이드 순서 조정 (원본 바로 뒤로 이동)
        if self.insert_mode == 'after_original':
            self._move_slide(prs, len(prs.slides) - 1, slide_idx + 1)

        logger.info(
            "슬라이드 %d 뒤에 재구성 슬라이드 삽입 완료 (도형 %d개 생성)",
            slide_idx + 1, shapes_after - shapes_before,
        )

    def _copy_text_elements(
        self,
        prs: Presentation,
        slide_idx: int,
        new_slide,
        analysis: SlideAnalysisResult,
    ) -> None:
        """원본 슬라이드의 제목/텍스트 요소를 새 슬라이드에 복사한다."""
        for element in analysis.elements:
            if element.element_type == 'text' and element.metadata.get('is_title'):
                # 제목 텍스트를 텍스트 박스로 추가
                from pptx.util import Pt
                from pptx.enum.text import PP_ALIGN
                from src.utils.color_utils import hex_to_rgb

                txbox = new_slide.shapes.add_textbox(
                    element.position['x'],
                    element.position['y'],
                    element.position['width'],
                    element.position['height'],
                )
                tf = txbox.text_frame
                p = tf.paragraphs[0]
                p.text = element.content or ''
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    run.font.size = Pt(24)
                    run.font.bold = True

    def _move_slide(self, prs: Presentation, from_idx: int, to_idx: int) -> None:
        """슬라이드를 from_idx에서 to_idx로 이동한다."""
        try:
            slides_list = prs.slides._sldIdLst

            slide_ids = list(slides_list)
            if from_idx < len(slide_ids) and to_idx <= len(slide_ids):
                slide_elem = slide_ids[from_idx]
                slides_list.remove(slide_elem)
                if to_idx >= len(list(slides_list)):
                    slides_list.append(slide_elem)
                else:
                    slides_list.insert(to_idx, slide_elem)
        except Exception as e:
            logger.warning("슬라이드 이동 실패: %s (순서가 변경되지 않을 수 있음)", e)

    def get_output_path(self, input_path: str, output_path: str | None = None) -> str:
        """출력 파일 경로를 결정한다."""
        if output_path:
            return output_path
        p = Path(input_path)
        return str(p.parent / f"{p.stem}{self.output_suffix}{p.suffix}")

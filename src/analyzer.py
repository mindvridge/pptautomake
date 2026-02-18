"""Module 1: SlideAnalyzer - PPTX 파싱 및 슬라이드별 요소 추출

python-pptx로 프레젠테이션 전체를 분석하여 각 슬라이드의
도형, 이미지, 표, 텍스트 등 모든 요소의 메타데이터를 수집한다.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Any

from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

from src.utils.xml_utils import (
    get_text_from_xml,
    get_fill_color_from_xml,
    get_position_from_xml,
    count_child_shapes,
)

logger = logging.getLogger(__name__)


@dataclass
class SlideElement:
    """슬라이드 내 개별 요소의 메타데이터를 담는 데이터 클래스"""
    element_type: str           # 'text', 'image', 'table', 'group_shape', 'chart', 'connector'
    position: dict              # {x, y, width, height} in EMU
    content: Any = None         # 텍스트 내용 또는 이미지 바이너리
    xml_element: Any = None     # 원본 XML 요소
    slide_index: int = 0
    z_order: int = 0            # 레이어 순서
    children: list = field(default_factory=list)  # 그룹 내 하위 요소
    metadata: dict = field(default_factory=dict)  # 추가 메타데이터

    @property
    def area(self) -> int:
        """요소의 면적 (EMU^2)"""
        return self.position.get('width', 0) * self.position.get('height', 0)


@dataclass
class SlideAnalysisResult:
    """슬라이드 분석 결과"""
    slide_index: int
    elements: list[SlideElement] = field(default_factory=list)
    slide_width: int = 0
    slide_height: int = 0

    @property
    def slide_area(self) -> int:
        return self.slide_width * self.slide_height

    def get_elements_by_type(self, element_type: str) -> list[SlideElement]:
        return [e for e in self.elements if e.element_type == element_type]


class SlideAnalyzer:
    """PPTX 파일을 분석하여 슬라이드별 요소를 추출하는 모듈"""

    def __init__(self):
        self.presentation = None
        self.slide_width = 0
        self.slide_height = 0

    def analyze(self, pptx_path: str) -> list[SlideAnalysisResult]:
        """프레젠테이션 전체를 분석하여 SlideAnalysisResult 리스트를 반환한다."""
        logger.info("PPTX 파일 분석 시작: %s", pptx_path)
        self.presentation = Presentation(pptx_path)
        self.slide_width = self.presentation.slide_width
        self.slide_height = self.presentation.slide_height

        results = []
        for idx, slide in enumerate(self.presentation.slides):
            result = self._analyze_slide(slide, idx)
            results.append(result)
            logger.info(
                "슬라이드 %d 분석 완료: %d개 요소 발견",
                idx + 1, len(result.elements),
            )

        logger.info("전체 분석 완료: %d개 슬라이드", len(results))
        return results

    def _analyze_slide(self, slide, slide_index: int) -> SlideAnalysisResult:
        """단일 슬라이드를 분석한다."""
        result = SlideAnalysisResult(
            slide_index=slide_index,
            slide_width=self.slide_width,
            slide_height=self.slide_height,
        )

        for z_order, shape in enumerate(slide.shapes):
            elements = self._extract_shape(shape, slide_index, z_order)
            result.elements.extend(elements)

        return result

    def _extract_shape(
        self, shape, slide_index: int, z_order: int
    ) -> list[SlideElement]:
        """도형 하나를 분석하여 SlideElement 리스트를 반환한다."""
        elements = []
        position = {
            'x': shape.left or 0,
            'y': shape.top or 0,
            'width': shape.width or 0,
            'height': shape.height or 0,
        }

        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            children = self._extract_grouped_shapes(shape, slide_index)
            element = SlideElement(
                element_type='group_shape',
                position=position,
                xml_element=shape._element,
                slide_index=slide_index,
                z_order=z_order,
                children=children,
                metadata={
                    'child_count': len(children),
                    'has_text': any(
                        c.element_type == 'text' or
                        (c.content and isinstance(c.content, str) and c.content.strip())
                        for c in children
                    ),
                },
            )
            elements.append(element)

        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_data = self._extract_image(shape)
            element = SlideElement(
                element_type='image',
                position=position,
                content=image_data.get('blob'),
                xml_element=shape._element,
                slide_index=slide_index,
                z_order=z_order,
                metadata={
                    'content_type': image_data.get('content_type', ''),
                    'filename': image_data.get('filename', ''),
                    'image_width': image_data.get('image_width', 0),
                    'image_height': image_data.get('image_height', 0),
                },
            )
            elements.append(element)

        elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            table_data = self._extract_table(shape)
            element = SlideElement(
                element_type='table',
                position=position,
                content=table_data,
                xml_element=shape._element,
                slide_index=slide_index,
                z_order=z_order,
                metadata={
                    'rows': len(table_data),
                    'cols': len(table_data[0]) if table_data else 0,
                },
            )
            elements.append(element)

        elif hasattr(shape, 'has_chart') and shape.has_chart:
            element = SlideElement(
                element_type='chart',
                position=position,
                xml_element=shape._element,
                slide_index=slide_index,
                z_order=z_order,
                metadata={'chart_type': str(shape.chart.chart_type) if shape.has_chart else ''},
            )
            elements.append(element)

        elif hasattr(shape, 'has_text_frame') and shape.has_text_frame:
            text = shape.text_frame.text.strip() if shape.text_frame else ''
            if text:
                element = SlideElement(
                    element_type='text',
                    position=position,
                    content=text,
                    xml_element=shape._element,
                    slide_index=slide_index,
                    z_order=z_order,
                    metadata={
                        'is_title': self._is_placeholder(shape),
                    },
                )
                elements.append(element)

        else:
            # 기타 도형 (자유형, 커넥터 등)
            element = SlideElement(
                element_type='shape',
                position=position,
                xml_element=shape._element,
                slide_index=slide_index,
                z_order=z_order,
            )
            elements.append(element)

        return elements

    @staticmethod
    def _is_placeholder(shape) -> bool:
        """도형이 플레이스홀더(제목, 부제 등)인지 확인한다."""
        try:
            return shape.placeholder_format is not None
        except (ValueError, AttributeError):
            return False

    def _extract_image(self, shape) -> dict:
        """이미지 도형에서 바이너리 데이터와 메타데이터를 추출한다."""
        try:
            image = shape.image
            return {
                'blob': image.blob,
                'content_type': image.content_type,
                'filename': getattr(image, 'filename', '') or '',
                'image_width': image.size[0] if hasattr(image, 'size') else 0,
                'image_height': image.size[1] if hasattr(image, 'size') else 0,
            }
        except Exception as e:
            logger.warning("이미지 추출 실패: %s", e)
            return {}

    def _extract_table(self, shape) -> list[list[str]]:
        """표 도형에서 셀 텍스트를 2D 리스트로 추출한다."""
        table = shape.table
        rows = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            rows.append(row_data)
        return rows

    def _extract_grouped_shapes(
        self, group_shape, slide_index: int
    ) -> list[SlideElement]:
        """그룹 도형을 재귀적으로 분석하여 하위 요소 리스트를 반환한다."""
        children = []
        try:
            for z_order, shape in enumerate(group_shape.shapes):
                child_elements = self._extract_shape(shape, slide_index, z_order)
                children.extend(child_elements)
        except Exception as e:
            logger.warning("그룹 도형 분석 실패: %s", e)
            # XML에서 직접 분석 시도
            xml_elem = group_shape._element
            child_count = count_child_shapes(xml_elem)
            if child_count > 0:
                text = get_text_from_xml(xml_elem)
                color = get_fill_color_from_xml(xml_elem)
                pos = get_position_from_xml(xml_elem) or {
                    'x': 0, 'y': 0, 'width': 0, 'height': 0
                }
                children.append(SlideElement(
                    element_type='shape',
                    position=pos,
                    content=text,
                    xml_element=xml_elem,
                    slide_index=slide_index,
                    metadata={
                        'fill_color': color,
                        'child_count': child_count,
                        'from_xml_fallback': True,
                    },
                ))
        return children

"""Module 2: ElementClassifier - 추출된 요소를 의미 단위로 분류

이미지가 사진인지 도식/다이어그램인지 판별하고,
그룹 도형의 레이아웃 패턴을 분석하여 DiagramType을 할당한다.
"""

from __future__ import annotations

import io
import logging
from enum import Enum
from dataclasses import dataclass, field

from PIL import Image
import numpy as np

from src.analyzer import SlideElement, SlideAnalysisResult
from src.utils.xml_utils import has_connector, count_child_shapes

logger = logging.getLogger(__name__)


class DiagramType(Enum):
    """도식 유형 열거형"""
    FLOWCHART = "flowchart"
    PROCESS = "process"
    COMPARISON_TABLE = "comparison_table"
    HIERARCHY = "hierarchy"
    CYCLE = "cycle"
    GRID = "grid"
    TIMELINE = "timeline"
    PIE_CHART = "pie_chart"
    BAR_CHART = "bar_chart"
    INFOGRAPHIC = "infographic"
    UNKNOWN = "unknown"


class ContentType(Enum):
    """콘텐츠 유형 열거형"""
    DIAGRAM_IMAGE = "diagram_image"       # 이미지로 삽입된 도식
    PHOTO = "photo"                       # 일반 사진/로고
    NATIVE_TABLE = "native_table"         # 네이티브 표
    NATIVE_CHART = "native_chart"         # 네이티브 차트
    GROUP_DIAGRAM = "group_diagram"       # 그룹 도형으로 된 도식
    TEXT = "text"                         # 텍스트
    DECORATIVE = "decorative"             # 장식용 요소 (아이콘, 배경)
    OTHER = "other"


@dataclass
class ClassifiedElement:
    """분류된 요소"""
    element: SlideElement
    content_type: ContentType
    diagram_type: DiagramType = DiagramType.UNKNOWN
    confidence: float = 0.0
    needs_rebuild: bool = False


@dataclass
class SlideClassificationResult:
    """슬라이드 분류 결과"""
    slide_index: int
    classified_elements: list[ClassifiedElement] = field(default_factory=list)
    needs_rebuild: bool = False

    @property
    def diagram_elements(self) -> list[ClassifiedElement]:
        return [
            ce for ce in self.classified_elements
            if ce.content_type in (ContentType.DIAGRAM_IMAGE, ContentType.GROUP_DIAGRAM)
        ]


class ElementClassifier:
    """추출된 요소를 의미 단위로 분류하는 모듈"""

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.min_diagram_area_ratio = self.config.get('min_diagram_area_ratio', 0.25)

    def classify_slide(
        self, analysis: SlideAnalysisResult
    ) -> SlideClassificationResult:
        """슬라이드 분석 결과를 분류한다."""
        result = SlideClassificationResult(slide_index=analysis.slide_index)

        for element in analysis.elements:
            classified = self._classify_element(element, analysis)
            result.classified_elements.append(classified)

        result.needs_rebuild = any(
            ce.needs_rebuild for ce in result.classified_elements
        )
        return result

    def classify_all(
        self, analyses: list[SlideAnalysisResult]
    ) -> list[SlideClassificationResult]:
        """모든 슬라이드 분석 결과를 분류한다."""
        results = []
        for analysis in analyses:
            result = self.classify_slide(analysis)
            results.append(result)
            logger.info(
                "슬라이드 %d 분류 완료: 재구성 필요=%s, 도식 요소=%d개",
                analysis.slide_index + 1,
                result.needs_rebuild,
                len(result.diagram_elements),
            )
        return results

    def _classify_element(
        self, element: SlideElement, analysis: SlideAnalysisResult
    ) -> ClassifiedElement:
        """단일 요소를 분류한다."""
        if element.element_type == 'image':
            return self._classify_image(element, analysis)
        elif element.element_type == 'group_shape':
            return self._classify_group_shape(element, analysis)
        elif element.element_type == 'table':
            return ClassifiedElement(
                element=element,
                content_type=ContentType.NATIVE_TABLE,
                confidence=1.0,
                needs_rebuild=False,
            )
        elif element.element_type == 'chart':
            return ClassifiedElement(
                element=element,
                content_type=ContentType.NATIVE_CHART,
                confidence=1.0,
                needs_rebuild=False,
            )
        elif element.element_type == 'text':
            return ClassifiedElement(
                element=element,
                content_type=ContentType.TEXT,
                confidence=1.0,
                needs_rebuild=False,
            )
        else:
            return ClassifiedElement(
                element=element,
                content_type=ContentType.OTHER,
                confidence=0.5,
                needs_rebuild=False,
            )

    def _classify_image(
        self, element: SlideElement, analysis: SlideAnalysisResult
    ) -> ClassifiedElement:
        """이미지 요소가 도식인지 사진인지 판별한다."""
        score = 0.0
        reasons = []

        # 1) 이미지 크기 - 슬라이드 면적의 일정 비율 이상이면 도식 가능성
        if analysis.slide_area > 0:
            area_ratio = element.area / analysis.slide_area
            if area_ratio >= self.min_diagram_area_ratio:
                score += 0.3
                reasons.append(f"면적 비율 {area_ratio:.1%}")
            elif area_ratio >= 0.15:
                score += 0.15
                reasons.append(f"중간 면적 {area_ratio:.1%}")

        # 2) 색상 분석
        if element.content:
            color_score = self._analyze_image_colors(element.content)
            score += color_score
            if color_score > 0:
                reasons.append("도식형 색상 패턴")

        # 3) 파일명 힌트
        filename = element.metadata.get('filename', '').lower()
        diagram_hints = ['그림', 'diagram', 'chart', 'graph', 'flow', 'process', '도식', '표']
        if any(hint in filename for hint in diagram_hints):
            score += 0.2
            reasons.append(f"파일명 힌트: {filename}")

        # 4) SVG는 주로 아이콘/장식용
        content_type = element.metadata.get('content_type', '')
        if 'svg' in content_type.lower():
            score -= 0.3
            reasons.append("SVG 아이콘 (장식용)")

        # 5) 너무 작은 이미지는 아이콘
        if element.area > 0 and analysis.slide_area > 0:
            if element.area / analysis.slide_area < 0.05:
                score -= 0.2
                reasons.append("소형 이미지 (아이콘)")

        is_diagram = score >= 0.3
        logger.debug(
            "이미지 분류: score=%.2f, is_diagram=%s, reasons=%s",
            score, is_diagram, reasons,
        )

        if is_diagram:
            return ClassifiedElement(
                element=element,
                content_type=ContentType.DIAGRAM_IMAGE,
                confidence=min(score, 1.0),
                needs_rebuild=True,
            )
        else:
            return ClassifiedElement(
                element=element,
                content_type=ContentType.PHOTO,
                confidence=max(1.0 - score, 0.0),
                needs_rebuild=False,
            )

    def _analyze_image_colors(self, image_blob: bytes) -> float:
        """이미지의 색상 패턴을 분석하여 도식 가능성 점수를 반환한다."""
        try:
            img = Image.open(io.BytesIO(image_blob))
            img = img.convert('RGB')
            img_small = img.resize((100, 100))
            pixels = np.array(img_small)

            # 고유 색상 수 계산 (도식은 제한된 색상 팔레트 사용)
            unique_colors = len(np.unique(
                pixels.reshape(-1, 3), axis=0
            ))

            # 도식은 보통 100개 이하의 고유 색상
            if unique_colors < 50:
                return 0.25
            elif unique_colors < 200:
                return 0.15
            elif unique_colors > 5000:
                return -0.1  # 사진일 가능성
            return 0.0
        except Exception as e:
            logger.debug("색상 분석 실패: %s", e)
            return 0.0

    def _classify_group_shape(
        self, element: SlideElement, analysis: SlideAnalysisResult
    ) -> ClassifiedElement:
        """그룹 도형을 분석하여 도식 유형을 분류한다."""
        child_count = len(element.children)
        has_text = element.metadata.get('has_text', False)
        xml_elem = element.xml_element
        has_connectors = has_connector(xml_elem) if xml_elem is not None else False

        # 하위 요소가 2개 이하면 장식용
        if child_count <= 2 and not has_text:
            return ClassifiedElement(
                element=element,
                content_type=ContentType.DECORATIVE,
                confidence=0.6,
                needs_rebuild=False,
            )

        diagram_type = self._detect_group_pattern(element, has_connectors)

        return ClassifiedElement(
            element=element,
            content_type=ContentType.GROUP_DIAGRAM,
            diagram_type=diagram_type,
            confidence=0.7,
            needs_rebuild=True,
        )

    def _detect_group_pattern(
        self, element: SlideElement, has_connectors: bool
    ) -> DiagramType:
        """그룹 도형의 배치 패턴을 분석하여 다이어그램 유형을 추론한다."""
        children = element.children
        if not children:
            return DiagramType.UNKNOWN

        # 커넥터가 있으면 플로우차트
        if has_connectors:
            return DiagramType.FLOWCHART

        # 하위 요소의 위치를 분석하여 패턴 추론
        positions = []
        for child in children:
            if child.position:
                positions.append(child.position)

        if len(positions) < 2:
            return DiagramType.UNKNOWN

        # x 좌표 기준으로 정렬
        x_coords = sorted(set(p['x'] for p in positions))
        y_coords = sorted(set(p['y'] for p in positions))

        # 2x2 패턴
        if len(x_coords) == 2 and len(y_coords) == 2:
            return DiagramType.GRID

        # 수평 배치 (y 좌표 유사)
        y_range = max(y_coords) - min(y_coords) if y_coords else 0
        x_range = max(x_coords) - min(x_coords) if x_coords else 0

        if y_range < x_range * 0.3 and len(positions) >= 3:
            return DiagramType.PROCESS

        # 수직 배치
        if x_range < y_range * 0.3 and len(positions) >= 3:
            return DiagramType.HIERARCHY

        # 그리드 (3x3 이상)
        if len(x_coords) >= 3 and len(y_coords) >= 3:
            return DiagramType.GRID

        return DiagramType.INFOGRAPHIC

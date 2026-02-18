"""ElementClassifier 테스트"""

import io
import pytest
from PIL import Image

from src.analyzer import SlideElement, SlideAnalysisResult
from src.classifier import (
    ElementClassifier,
    ContentType,
    DiagramType,
    ClassifiedElement,
)


def _make_analysis(elements: list[SlideElement], slide_width=12192000, slide_height=6858000):
    """테스트용 SlideAnalysisResult를 생성한다."""
    return SlideAnalysisResult(
        slide_index=0,
        elements=elements,
        slide_width=slide_width,
        slide_height=slide_height,
    )


def _create_simple_image(width=200, height=200, color=(255, 0, 0)):
    """단색 테스트 이미지를 생성한다."""
    img = Image.new('RGB', (width, height), color)
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()


def test_text_classified_correctly():
    """텍스트 요소가 TEXT로 분류되는지 확인"""
    element = SlideElement(
        element_type='text',
        position={'x': 0, 'y': 0, 'width': 1000, 'height': 500},
        content="테스트 텍스트",
    )
    analysis = _make_analysis([element])
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    assert len(result.classified_elements) == 1
    assert result.classified_elements[0].content_type == ContentType.TEXT
    assert result.needs_rebuild is False


def test_table_classified_correctly():
    """표 요소가 NATIVE_TABLE로 분류되는지 확인"""
    element = SlideElement(
        element_type='table',
        position={'x': 0, 'y': 0, 'width': 5000000, 'height': 3000000},
        content=[["A", "B"], ["1", "2"]],
    )
    analysis = _make_analysis([element])
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    assert result.classified_elements[0].content_type == ContentType.NATIVE_TABLE
    assert result.needs_rebuild is False


def test_large_image_classified_as_diagram():
    """슬라이드 면적 30% 이상의 단색 이미지가 도식으로 분류되는지 확인"""
    slide_w = 12192000
    slide_h = 6858000
    # 슬라이드 면적의 40% 차지하는 이미지
    img_w = int(slide_w * 0.7)
    img_h = int(slide_h * 0.6)

    image_blob = _create_simple_image(200, 200, (100, 100, 200))

    element = SlideElement(
        element_type='image',
        position={'x': 0, 'y': 0, 'width': img_w, 'height': img_h},
        content=image_blob,
        metadata={'content_type': 'image/png', 'filename': 'diagram.png'},
    )
    analysis = _make_analysis([element], slide_w, slide_h)
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    assert result.classified_elements[0].content_type == ContentType.DIAGRAM_IMAGE
    assert result.needs_rebuild is True


def test_small_image_not_diagram():
    """아이콘 크기 이미지가 도식으로 분류되지 않는지 확인"""
    slide_w = 12192000
    slide_h = 6858000
    # 슬라이드 면적의 2% (아이콘)
    img_w = int(slide_w * 0.1)
    img_h = int(slide_h * 0.1)

    element = SlideElement(
        element_type='image',
        position={'x': 0, 'y': 0, 'width': img_w, 'height': img_h},
        content=_create_simple_image(50, 50),
        metadata={'content_type': 'image/png', 'filename': 'icon.png'},
    )
    analysis = _make_analysis([element], slide_w, slide_h)
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    assert result.classified_elements[0].content_type == ContentType.PHOTO


def test_group_shape_with_children_is_diagram():
    """하위 요소가 3개 이상인 그룹 도형이 도식으로 분류되는지 확인"""
    children = [
        SlideElement(
            element_type='text',
            position={'x': i * 1000000, 'y': 1000000, 'width': 800000, 'height': 400000},
            content=f"Step {i+1}",
        )
        for i in range(4)
    ]

    element = SlideElement(
        element_type='group_shape',
        position={'x': 0, 'y': 0, 'width': 10000000, 'height': 5000000},
        children=children,
        metadata={'child_count': 4, 'has_text': True},
    )
    analysis = _make_analysis([element])
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    assert result.classified_elements[0].content_type == ContentType.GROUP_DIAGRAM
    assert result.needs_rebuild is True


def test_group_shape_horizontal_pattern():
    """수평 배치 그룹이 PROCESS로 분류되는지 확인"""
    children = [
        SlideElement(
            element_type='text',
            position={'x': i * 2000000, 'y': 1000000, 'width': 1500000, 'height': 1000000},
            content=f"Step {i+1}",
        )
        for i in range(5)
    ]

    element = SlideElement(
        element_type='group_shape',
        position={'x': 0, 'y': 0, 'width': 12000000, 'height': 3000000},
        children=children,
        metadata={'child_count': 5, 'has_text': True},
    )
    analysis = _make_analysis([element])
    classifier = ElementClassifier()
    result = classifier.classify_slide(analysis)

    ce = result.classified_elements[0]
    assert ce.diagram_type == DiagramType.PROCESS


def test_classify_all():
    """classify_all이 여러 슬라이드를 올바르게 처리하는지 확인"""
    analyses = [
        _make_analysis([
            SlideElement(
                element_type='text',
                position={'x': 0, 'y': 0, 'width': 1000, 'height': 500},
                content="Slide 1",
            )
        ]),
        _make_analysis([
            SlideElement(
                element_type='text',
                position={'x': 0, 'y': 0, 'width': 1000, 'height': 500},
                content="Slide 2",
            )
        ]),
    ]
    analyses[1].slide_index = 1

    classifier = ElementClassifier()
    results = classifier.classify_all(analyses)

    assert len(results) == 2

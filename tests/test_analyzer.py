"""SlideAnalyzer 테스트"""

import pytest
from pptx import Presentation
from pptx.util import Inches

from src.analyzer import SlideAnalyzer, SlideElement, SlideAnalysisResult


@pytest.fixture
def sample_pptx(tmp_path):
    """테스트용 PPTX 파일을 생성한다."""
    pptx_path = tmp_path / "test.pptx"
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 슬라이드 1: 텍스트만
    slide_layout = prs.slide_layouts[5]  # 빈 레이아웃
    slide = prs.slides.add_slide(slide_layout)
    txbox = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(5), Inches(1),
    )
    txbox.text_frame.text = "테스트 제목"

    # 슬라이드 2: 표 포함
    slide2 = prs.slides.add_slide(slide_layout)
    table_shape = slide2.shapes.add_table(
        3, 2,
        Inches(1), Inches(1), Inches(8), Inches(3),
    )
    table = table_shape.table
    table.cell(0, 0).text = "항목"
    table.cell(0, 1).text = "값"
    table.cell(1, 0).text = "A"
    table.cell(1, 1).text = "100"
    table.cell(2, 0).text = "B"
    table.cell(2, 1).text = "200"

    prs.save(str(pptx_path))
    return str(pptx_path)


def test_analyze_returns_results(sample_pptx):
    """분석이 SlideAnalysisResult 리스트를 반환하는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    assert isinstance(results, list)
    assert len(results) == 2
    assert all(isinstance(r, SlideAnalysisResult) for r in results)


def test_analyze_slide_indices(sample_pptx):
    """슬라이드 인덱스가 올바르게 할당되는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    assert results[0].slide_index == 0
    assert results[1].slide_index == 1


def test_text_element_extraction(sample_pptx):
    """텍스트 요소가 올바르게 추출되는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    text_elements = results[0].get_elements_by_type('text')
    assert len(text_elements) >= 1
    assert any("테스트 제목" in (e.content or '') for e in text_elements)


def test_table_element_extraction(sample_pptx):
    """표 요소가 올바르게 추출되는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    table_elements = results[1].get_elements_by_type('table')
    assert len(table_elements) == 1

    table_data = table_elements[0].content
    assert table_data is not None
    assert len(table_data) == 3  # 3행
    assert len(table_data[0]) == 2  # 2열
    assert table_data[0][0] == "항목"
    assert table_data[1][1] == "100"


def test_slide_dimensions(sample_pptx):
    """슬라이드 크기가 올바르게 기록되는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    assert results[0].slide_width > 0
    assert results[0].slide_height > 0
    assert results[0].slide_area > 0


def test_element_position(sample_pptx):
    """요소 위치 정보가 올바르게 추출되는지 확인"""
    analyzer = SlideAnalyzer()
    results = analyzer.analyze(sample_pptx)

    for element in results[0].elements:
        assert 'x' in element.position
        assert 'y' in element.position
        assert 'width' in element.position
        assert 'height' in element.position


def test_slide_element_area():
    """SlideElement.area 속성이 올바르게 계산되는지 확인"""
    element = SlideElement(
        element_type='image',
        position={'x': 0, 'y': 0, 'width': 1000, 'height': 500},
    )
    assert element.area == 500000

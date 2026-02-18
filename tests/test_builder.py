"""NativePPTXBuilder 테스트"""

import pytest
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE_TYPE

from src.builder import NativePPTXBuilder
from src.diagram_analyzer import DiagramData


@pytest.fixture
def builder():
    """기본 설정의 빌더 인스턴스"""
    return NativePPTXBuilder()


@pytest.fixture
def blank_slide():
    """빈 슬라이드를 포함한 프레젠테이션"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide_layout = prs.slide_layouts[6]  # 빈 레이아웃
    slide = prs.slides.add_slide(slide_layout)
    return prs, slide


def _make_process_data(node_count=5):
    """프로세스 다이어그램 테스트 데이터"""
    nodes = [
        {
            'id': f'node{i}',
            'text': f'Step {i+1}',
            'color_fill': '#4A90D9',
            'color_text': '#FFFFFF',
            'shape': 'rounded_rect',
            'relative_position': {'row': 0, 'col': i},
            'relative_size': 'medium',
        }
        for i in range(node_count)
    ]
    return DiagramData(
        diagram_type='process',
        title='프로세스 테스트',
        nodes=nodes,
        layout={'direction': 'horizontal', 'rows': 1, 'cols': node_count, 'spacing': 'even'},
        style={'corner_radius': 'rounded', 'has_shadow': False},
    )


def test_process_diagram_node_count(builder, blank_slide):
    """5단계 프로세스에서 노드가 정확히 생성되는지 확인"""
    prs, slide = blank_slide
    data = _make_process_data(5)
    builder.build_process_diagram(slide, data)

    # 텍스트 박스(제목) + 5 노드 + 4 화살표 = 10
    shapes = list(slide.shapes)
    # 제목 1 + 노드 5 + 화살표 4 = 10
    assert len(shapes) == 10


def test_process_diagram_no_nodes(builder, blank_slide):
    """노드가 없는 경우 에러 없이 종료되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(diagram_type='process', nodes=[])
    builder.build_process_diagram(slide, data)

    assert len(list(slide.shapes)) == 0


def test_comparison_table_dimensions(builder, blank_slide):
    """표가 올바른 차원으로 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='comparison_table',
        title='비교표 테스트',
        table_data=[
            ['항목', '값A', '값B'],
            ['속도', '빠름', '보통'],
            ['가격', '높음', '낮음'],
            ['품질', '상', '중'],
            ['지원', 'O', 'X'],
        ],
    )
    builder.build_comparison_table(slide, data)

    # 제목 + 표 = 2
    shapes = list(slide.shapes)
    assert len(shapes) == 2

    # 표 찾기
    table_shape = None
    for shape in shapes:
        if shape.has_table:
            table_shape = shape
            break

    assert table_shape is not None
    assert len(table_shape.table.rows) == 5
    assert len(table_shape.table.columns) == 3


def test_grid_layout(builder, blank_slide):
    """그리드 레이아웃이 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='grid',
        title='그리드 테스트',
        nodes=[
            {'id': f'node{i}', 'text': f'Cell {i+1}', 'color_fill': '#4A90D9', 'color_text': '#FFFFFF'}
            for i in range(4)
        ],
        layout={'direction': 'grid', 'rows': 2, 'cols': 2},
    )
    builder.build_grid_layout(slide, data)

    # 제목 1 + 셀 4 = 5
    shapes = list(slide.shapes)
    assert len(shapes) == 5


def test_build_dispatches_correctly(builder, blank_slide):
    """build()가 diagram_type에 따라 올바른 빌더를 호출하는지 확인"""
    prs, slide = blank_slide
    data = _make_process_data(3)
    builder.build(slide, data)

    # 빌드가 정상적으로 수행되었는지 확인
    assert len(list(slide.shapes)) > 0


def test_native_elements_no_images(builder, blank_slide):
    """재구성된 슬라이드에 이미지 요소가 없는지 확인 (모두 네이티브)"""
    prs, slide = blank_slide
    data = _make_process_data(3)
    builder.build(slide, data)

    for shape in slide.shapes:
        assert shape.shape_type != MSO_SHAPE_TYPE.PICTURE


def test_timeline_creation(builder, blank_slide):
    """타임라인이 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='timeline',
        title='타임라인 테스트',
        nodes=[
            {'id': f'node{i}', 'text': f'Phase {i+1}', 'color_fill': '#2D6A4F', 'color_text': '#FFFFFF'}
            for i in range(4)
        ],
    )
    builder.build_timeline(slide, data)

    # 제목 + 기준선 + (마커 + 텍스트) * 4 = 10
    shapes = list(slide.shapes)
    assert len(shapes) >= 9  # 최소 제목 + 선 + 4마커 + 4텍스트


def test_bar_chart_creation(builder, blank_slide):
    """막대 차트가 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='bar_chart',
        title='차트 테스트',
        chart_data={
            'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
            'series': [
                {'name': '매출', 'values': [100, 150, 130, 200]},
            ],
        },
    )
    builder.build_bar_chart(slide, data)

    # 제목 + 차트 = 2
    shapes = list(slide.shapes)
    assert len(shapes) == 2

    # 차트 확인
    chart_shape = None
    for shape in shapes:
        if shape.has_chart:
            chart_shape = shape
            break
    assert chart_shape is not None


def test_pie_chart_creation(builder, blank_slide):
    """원형 차트가 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='pie_chart',
        title='파이 차트 테스트',
        chart_data={
            'categories': ['A', 'B', 'C'],
            'series': [
                {'name': '비율', 'values': [40, 35, 25]},
            ],
        },
    )
    builder.build_pie_chart(slide, data)

    chart_found = False
    for shape in slide.shapes:
        if shape.has_chart:
            chart_found = True
            break
    assert chart_found


def test_vertical_flowchart(builder, blank_slide):
    """수직 플로우차트가 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='flowchart',
        title='플로우차트 테스트',
        nodes=[
            {'id': 'node0', 'text': '시작', 'color_fill': '#4A90D9', 'color_text': '#FFFFFF'},
            {'id': 'node1', 'text': '처리', 'color_fill': '#4A90D9', 'color_text': '#FFFFFF'},
            {'id': 'node2', 'text': '종료', 'color_fill': '#4A90D9', 'color_text': '#FFFFFF'},
        ],
        layout={'direction': 'vertical'},
    )
    builder.build_flowchart(slide, data)

    # 노드와 화살표가 생성되었는지 확인
    assert len(list(slide.shapes)) > 0


def test_cycle_diagram(builder, blank_slide):
    """순환 다이어그램이 올바르게 생성되는지 확인"""
    prs, slide = blank_slide
    data = DiagramData(
        diagram_type='cycle',
        title='순환 테스트',
        nodes=[
            {'id': f'node{i}', 'text': f'Phase {i+1}', 'color_fill': '#4A90D9', 'color_text': '#FFFFFF'}
            for i in range(4)
        ],
    )
    builder.build_cycle_diagram(slide, data)

    assert len(list(slide.shapes)) > 0

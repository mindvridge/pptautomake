"""Module 4: NativePPTXBuilder - 네이티브 PPTX 요소 생성

분석된 도식 데이터를 python-pptx 네이티브 요소(도형, 표, 차트, 커넥터)로 변환한다.
이것이 플러그인의 핵심 모듈이다.
"""

from __future__ import annotations

import logging
import math

from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor

from src.diagram_analyzer import DiagramData
from src.utils.color_utils import hex_to_rgb, get_contrast_text_color
from src.utils.layout_utils import (
    calculate_even_spacing,
    calculate_grid_positions,
    calculate_circular_positions,
    DEFAULT_SLIDE_WIDTH,
    DEFAULT_SLIDE_HEIGHT,
)

logger = logging.getLogger(__name__)


class NativePPTXBuilder:
    """분석된 도식 데이터를 네이티브 PPTX 요소로 변환하는 빌더"""

    # 빌더 메서드 매핑
    BUILDER_MAP = {
        'process': 'build_process_diagram',
        'flowchart': 'build_flowchart',
        'comparison_table': 'build_comparison_table',
        'grid': 'build_grid_layout',
        'timeline': 'build_timeline',
        'cycle': 'build_cycle_diagram',
        'bar_chart': 'build_bar_chart',
        'pie_chart': 'build_pie_chart',
        'hierarchy': 'build_hierarchy',
        'infographic': 'build_infographic',
    }

    # 노드 텍스트 최대 길이 (넘치면 잘라냄)
    MAX_NODE_TEXT_LEN = 200

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.slide_width = Inches(self.config.get('slide_width_inches', 13.333))
        self.slide_height = Inches(self.config.get('slide_height_inches', 7.5))
        self.margin = Inches(self.config.get('margin_inches', 0.5))

        colors = self.config.get('default_colors', {})
        self.color_primary = colors.get('primary', '#2D6A4F')
        self.color_secondary = colors.get('secondary', '#40916C')
        self.color_accent = colors.get('accent', '#95D5B2')
        self.color_bg = colors.get('background', '#F0F0F0')
        self.color_text = colors.get('text', '#333333')

        fonts = self.config.get('default_fonts', {})
        self.font_title = fonts.get('title', '맑은 고딕')
        self.font_body = fonts.get('body', '맑은 고딕')
        self.title_size = Pt(fonts.get('title_size_pt', 24))
        self.body_size = Pt(fonts.get('body_size_pt', 12))

    def build(self, slide, diagram_data: DiagramData) -> None:
        """도식 데이터에 맞는 빌더 메서드를 호출하여 슬라이드에 요소를 생성한다."""
        diagram_type = diagram_data.diagram_type
        builder_name = self.BUILDER_MAP.get(diagram_type, 'build_process_diagram')
        builder_method = getattr(self, builder_name, self.build_process_diagram)

        logger.info("빌더 실행: %s (유형: %s)", builder_name, diagram_type)

        try:
            builder_method(slide, diagram_data)
        except Exception as e:
            logger.error("빌더 실행 실패 (%s): %s", builder_name, e)
            # 폴백: 프로세스 다이어그램으로 시도
            if builder_name != 'build_process_diagram':
                logger.info("폴백: build_process_diagram 시도")
                self.build_process_diagram(slide, diagram_data)

    def build_process_diagram(self, slide, diagram_data: DiagramData) -> None:
        """프로세스 단계형 다이어그램을 생성한다.

        각 단계를 둥근 사각형으로 생성하고 단계 사이에 화살표 커넥터를 추가한다.
        """
        nodes = diagram_data.nodes
        if not nodes:
            logger.warning("노드가 없어 프로세스 다이어그램을 생성할 수 없습니다")
            return

        # 제목 추가
        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        node_count = len(nodes)
        arrow_width = Inches(0.4)
        arrow_count = node_count - 1
        spacing = Inches(0.2)
        available_width = self.slide_width - 2 * self.margin
        total_arrows = arrow_count * (arrow_width + spacing)
        node_width = (available_width - total_arrows - (node_count - 1) * spacing) / node_count
        node_height = Inches(2.0)
        top = Inches(2.5)

        for i, node in enumerate(nodes):
            # 노드 위치 계산
            x = self.margin + i * (node_width + arrow_width + spacing * 2)

            # 둥근 사각형 생성
            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(x), int(top), int(node_width), int(node_height),
            )

            # 채우기 색상
            fill_color = node.get('color_fill', self.color_primary)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_to_rgb(fill_color)

            # 테두리
            shape.line.fill.background()

            # 텍스트
            tf = shape.text_frame
            tf.word_wrap = True
            tf.auto_size = None

            # 기존 단락의 텍스트 설정
            p = tf.paragraphs[0]
            p.text = self._get_node_text(node)
            p.alignment = PP_ALIGN.CENTER

            text_color = node.get('color_text', get_contrast_text_color(fill_color))
            for run in p.runs:
                run.font.size = self.body_size
                run.font.name = self.font_body
                run.font.color.rgb = hex_to_rgb(text_color)

            # 텍스트 수직 정렬
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            shape.text_frame.paragraphs[0].space_before = Pt(0)
            shape.text_frame.paragraphs[0].space_after = Pt(0)

            # 화살표 커넥터 (마지막 노드 제외)
            if i < node_count - 1:
                arrow_x = int(x + node_width + spacing)
                arrow_y = int(top + node_height / 2 - Inches(0.15))
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.RIGHT_ARROW,
                    arrow_x, arrow_y,
                    int(arrow_width), int(Inches(0.3)),
                )
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = hex_to_rgb(self.color_text)
                arrow.line.fill.background()

    def build_comparison_table(self, slide, diagram_data: DiagramData) -> None:
        """비교표를 네이티브 표로 생성한다."""
        # 표 데이터 결정
        if diagram_data.table_data:
            table_data = diagram_data.table_data
        elif diagram_data.nodes:
            # 노드 데이터를 표 형태로 변환
            table_data = self._nodes_to_table_data(diagram_data)
        else:
            logger.warning("표 데이터가 없습니다")
            return

        if not table_data or not table_data[0]:
            return

        # 제목 추가
        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        rows = len(table_data)
        cols = len(table_data[0])
        table_width = self.slide_width - 2 * self.margin
        table_height = Inches(min(rows * 0.6, 5.0))
        top = Inches(1.8)

        table_shape = slide.shapes.add_table(
            rows, cols,
            int(self.margin), int(top),
            int(table_width), int(table_height),
        )
        table = table_shape.table

        for r, row_data in enumerate(table_data):
            for c, cell_text in enumerate(row_data):
                cell = table.cell(r, c)
                cell.text = str(cell_text)

                # 셀 스타일링
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER

                if p.runs:
                    for run in p.runs:
                        run.font.size = self.body_size
                        run.font.name = self.font_body

                # 헤더 행 스타일링
                if r == 0:
                    self._style_table_header_cell(cell)
                else:
                    # 데이터 행 교대 색상
                    if r % 2 == 0:
                        self._set_cell_fill(cell, '#F5F5F5')

    def build_flowchart(self, slide, diagram_data: DiagramData) -> None:
        """플로우차트를 생성한다. 노드 + 커넥터 기반."""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        direction = diagram_data.layout.get('direction', 'horizontal')
        is_vertical = direction == 'vertical'

        if is_vertical:
            self._build_vertical_flowchart(slide, diagram_data)
        else:
            # 수평 배치는 프로세스 다이어그램과 유사
            self.build_process_diagram(slide, diagram_data)

    def _build_vertical_flowchart(self, slide, diagram_data: DiagramData) -> None:
        """수직 플로우차트를 생성한다."""
        nodes = diagram_data.nodes
        node_count = len(nodes)
        node_width = Inches(4.0)
        node_height = Inches(1.0)
        spacing = Inches(0.5)
        arrow_height = Inches(0.3)
        center_x = int(self.slide_width / 2 - node_width / 2)
        start_y = Inches(1.5)

        for i, node in enumerate(nodes):
            y = int(start_y + i * (node_height + arrow_height + spacing))

            shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                center_x, y, int(node_width), int(node_height),
            )

            fill_color = node.get('color_fill', self.color_primary)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
            shape.line.fill.background()

            tf = shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = self._get_node_text(node)
            p.alignment = PP_ALIGN.CENTER
            text_color = node.get('color_text', get_contrast_text_color(fill_color))
            for run in p.runs:
                run.font.size = self.body_size
                run.font.name = self.font_body
                run.font.color.rgb = hex_to_rgb(text_color)

            # 아래쪽 화살표 (마지막 노드 제외)
            if i < node_count - 1:
                arrow_x = int(self.slide_width / 2 - Inches(0.15))
                arrow_y = int(y + node_height + spacing / 4)
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW,
                    arrow_x, arrow_y,
                    int(Inches(0.3)), int(arrow_height),
                )
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = hex_to_rgb(self.color_text)
                arrow.line.fill.background()

    def build_grid_layout(self, slide, diagram_data: DiagramData) -> None:
        """2x2 / 3x3 그리드 레이아웃을 생성한다."""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        layout = diagram_data.layout
        rows = layout.get('rows', 2)
        cols = layout.get('cols', 2)

        # 노드 수에 맞게 그리드 크기 조정
        node_count = len(nodes)
        if rows * cols < node_count:
            cols = math.ceil(math.sqrt(node_count))
            rows = math.ceil(node_count / cols)

        grid = calculate_grid_positions(
            rows, cols,
            slide_width=self.slide_width,
            slide_height=self.slide_height,
            margin=self.margin,
        )

        node_idx = 0
        for r in range(rows):
            for c in range(cols):
                if node_idx >= len(nodes):
                    break
                node = nodes[node_idx]
                pos = grid[r][c]

                shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    pos['x'], pos['y'], pos['width'], pos['height'],
                )

                fill_color = node.get('color_fill', self.color_primary)
                shape.fill.solid()
                shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
                shape.line.fill.background()

                tf = shape.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = self._get_node_text(node)
                p.alignment = PP_ALIGN.CENTER
                text_color = node.get('color_text', get_contrast_text_color(fill_color))
                for run in p.runs:
                    run.font.size = self.body_size
                    run.font.name = self.font_body
                    run.font.color.rgb = hex_to_rgb(text_color)

                node_idx += 1

    def build_timeline(self, slide, diagram_data: DiagramData) -> None:
        """타임라인/단계 다이어그램을 생성한다."""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        node_count = len(nodes)
        line_y = int(Inches(4.0))
        line_height = int(Inches(0.05))

        # 수평 기준선
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(self.margin), line_y,
            int(self.slide_width - 2 * self.margin), line_height,
        )
        line.fill.solid()
        line.fill.fore_color.rgb = hex_to_rgb(self.color_text)
        line.line.fill.background()

        # 마커 배치
        positions = calculate_even_spacing(
            self.slide_width, node_count, margin=self.margin
        )

        for i, (node, pos) in enumerate(zip(nodes, positions)):
            # 마커 원
            marker_size = int(Inches(0.4))
            marker_x = pos['x'] + pos['width'] // 2 - marker_size // 2
            marker_y = line_y - marker_size // 2

            marker = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                marker_x, marker_y, marker_size, marker_size,
            )
            fill_color = node.get('color_fill', self.color_primary)
            marker.fill.solid()
            marker.fill.fore_color.rgb = hex_to_rgb(fill_color)
            marker.line.fill.background()

            # 텍스트 박스 (마커 위/아래 교대)
            text_width = int(pos['width'])
            text_height = int(Inches(1.2))
            text_x = pos['x']

            if i % 2 == 0:
                text_y = line_y - marker_size - text_height - int(Inches(0.1))
            else:
                text_y = line_y + marker_size + int(Inches(0.1))

            txbox = slide.shapes.add_textbox(
                text_x, text_y, text_width, text_height,
            )
            tf = txbox.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = self._get_node_text(node)
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.name = self.font_body
                run.font.color.rgb = hex_to_rgb(self.color_text)

    def build_cycle_diagram(self, slide, diagram_data: DiagramData) -> None:
        """순환 다이어그램을 생성한다. (원형 배치)"""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        node_count = len(nodes)
        center_x = self.slide_width // 2
        center_y = int(Inches(4.2))
        radius = Inches(2.2)
        node_size = Inches(1.5)

        positions = calculate_circular_positions(
            node_count, center_x=center_x, center_y=center_y, radius=radius
        )

        for i, (node, pos) in enumerate(zip(nodes, positions)):
            x = pos['x'] - int(node_size / 2)
            y = pos['y'] - int(node_size / 2)

            shape = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                x, y, int(node_size), int(node_size),
            )

            fill_color = node.get('color_fill', self.color_primary)
            shape.fill.solid()
            shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
            shape.line.fill.background()

            tf = shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = self._get_node_text(node)
            p.alignment = PP_ALIGN.CENTER
            text_color = node.get('color_text', get_contrast_text_color(fill_color))
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.name = self.font_body
                run.font.color.rgb = hex_to_rgb(text_color)

            # 다음 노드로의 화살표 (곡선 대신 직선 화살표)
            if node_count > 1:
                next_idx = (i + 1) % node_count
                next_pos = positions[next_idx]

                # 화살표 시작/끝 점 계산 (원 가장자리)
                dx = next_pos['x'] - pos['x']
                dy = next_pos['y'] - pos['y']
                dist = math.sqrt(dx * dx + dy * dy)
                if dist > 0:
                    # 원 가장자리에서 시작
                    offset = int(node_size / 2)
                    start_x = pos['x'] + int(dx / dist * offset)
                    start_y = pos['y'] + int(dy / dist * offset)
                    end_x = next_pos['x'] - int(dx / dist * offset)
                    end_y = next_pos['y'] - int(dy / dist * offset)

                    arrow_len = int(math.sqrt(
                        (end_x - start_x) ** 2 + (end_y - start_y) ** 2
                    ))
                    if arrow_len > 0:
                        # 화살표 도형으로 표현
                        arrow = slide.shapes.add_shape(
                            MSO_SHAPE.RIGHT_ARROW,
                            min(start_x, end_x),
                            min(start_y, end_y),
                            max(abs(end_x - start_x), Inches(0.2)),
                            max(abs(end_y - start_y), Inches(0.2)),
                        )
                        arrow.fill.solid()
                        arrow.fill.fore_color.rgb = hex_to_rgb(self.color_text)
                        arrow.line.fill.background()
                        arrow.rotation = math.degrees(math.atan2(dy, dx))

    def build_bar_chart(self, slide, diagram_data: DiagramData) -> None:
        """막대 차트를 네이티브 Chart로 생성한다."""
        chart_data = diagram_data.chart_data
        if not chart_data and diagram_data.nodes:
            # 노드 데이터에서 차트 데이터 추출 시도
            chart_data = self._nodes_to_chart_data(diagram_data)

        if not chart_data:
            logger.warning("차트 데이터가 없습니다")
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = chart_data.get('categories', [])
        for series in chart_data.get('series', []):
            chart_data_obj.add_series(
                series.get('name', ''),
                series.get('values', []),
            )

        chart_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            int(Inches(1)), int(Inches(2)),
            int(Inches(10)), int(Inches(5)),
            chart_data_obj,
        )

        chart = chart_frame.chart
        chart.has_legend = True
        if chart_data.get('series') and len(chart_data['series']) == 1:
            chart.has_legend = False

    def build_pie_chart(self, slide, diagram_data: DiagramData) -> None:
        """원형 차트를 네이티브 Chart로 생성한다."""
        chart_data = diagram_data.chart_data
        if not chart_data and diagram_data.nodes:
            chart_data = self._nodes_to_chart_data(diagram_data)

        if not chart_data:
            logger.warning("차트 데이터가 없습니다")
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        chart_data_obj = CategoryChartData()
        chart_data_obj.categories = chart_data.get('categories', [])
        for series in chart_data.get('series', []):
            chart_data_obj.add_series(
                series.get('name', ''),
                series.get('values', []),
            )

        chart_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE,
            int(Inches(2)), int(Inches(2)),
            int(Inches(8)), int(Inches(5)),
            chart_data_obj,
        )

        chart = chart_frame.chart
        chart.has_legend = True

    def build_hierarchy(self, slide, diagram_data: DiagramData) -> None:
        """계층 구조 다이어그램을 생성한다. (트리형 상하 배치)"""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        connections = diagram_data.connections

        # 레벨별 노드 분류 (connections에서 부모-자식 관계 추론)
        node_map = {n.get('id', f'node{i}'): n for i, n in enumerate(nodes)}
        children_of: dict[str, list[str]] = {}
        has_parent: set[str] = set()

        for conn in connections:
            parent = conn.get('from', '')
            child = conn.get('to', '')
            if parent and child:
                children_of.setdefault(parent, []).append(child)
                has_parent.add(child)

        # 루트 노드 찾기 (부모가 없는 노드)
        roots = [nid for nid in node_map if nid not in has_parent]
        if not roots:
            roots = [nodes[0].get('id', 'node0')]

        # BFS로 레벨 할당
        levels: list[list[str]] = []
        visited = set()
        queue = list(roots)
        while queue:
            levels.append(queue)
            visited.update(queue)
            next_level = []
            for nid in queue:
                for child in children_of.get(nid, []):
                    if child not in visited:
                        next_level.append(child)
            queue = next_level

        # connections가 없으면 단순 수직 배치
        if len(levels) == 1 and len(nodes) > 1:
            self._build_vertical_flowchart(slide, diagram_data)
            return

        # 레벨별 렌더링
        available_width = self.slide_width - 2 * self.margin
        level_height = Inches(1.2)
        spacing_y = Inches(0.6)
        start_y = Inches(1.5)

        for level_idx, level_nodes in enumerate(levels):
            y = int(start_y + level_idx * (level_height + spacing_y))
            count = len(level_nodes)
            if count == 0:
                continue

            node_width = min(
                int(available_width / count - Inches(0.3)),
                int(Inches(3.5)),
            )
            total_width = count * node_width + (count - 1) * int(Inches(0.3))
            start_x = int(self.slide_width / 2 - total_width / 2)

            for i, nid in enumerate(level_nodes):
                node = node_map.get(nid, {'text': nid, 'color_fill': self.color_primary})
                x = start_x + i * (node_width + int(Inches(0.3)))

                shape = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    x, y, node_width, int(level_height),
                )

                fill_color = node.get('color_fill', self.color_primary)
                shape.fill.solid()
                shape.fill.fore_color.rgb = hex_to_rgb(fill_color)
                shape.line.fill.background()

                tf = shape.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = self._get_node_text(node)
                p.alignment = PP_ALIGN.CENTER
                text_color = node.get('color_text', get_contrast_text_color(fill_color))
                for run in p.runs:
                    run.font.size = self.body_size
                    run.font.name = self.font_body
                    run.font.color.rgb = hex_to_rgb(text_color)

            # 레벨 간 연결선 (아래쪽 화살표)
            if level_idx < len(levels) - 1:
                arrow_y = int(y + level_height + int(spacing_y * 0.1))
                arrow_x = int(self.slide_width / 2 - Inches(0.15))
                arrow = slide.shapes.add_shape(
                    MSO_SHAPE.DOWN_ARROW,
                    arrow_x, arrow_y,
                    int(Inches(0.3)), int(spacing_y * 0.6),
                )
                arrow.fill.solid()
                arrow.fill.fore_color.rgb = hex_to_rgb(self.color_text)
                arrow.line.fill.background()

    def build_infographic(self, slide, diagram_data: DiagramData) -> None:
        """인포그래픽 레이아웃을 생성한다. (아이콘 + 텍스트 카드)"""
        nodes = diagram_data.nodes
        if not nodes:
            return

        if diagram_data.title:
            self._add_title(slide, diagram_data.title)

        node_count = len(nodes)
        cols = min(node_count, 4)
        rows = math.ceil(node_count / cols)

        available_width = self.slide_width - 2 * self.margin
        available_height = self.slide_height - Inches(2.5)
        card_width = int(available_width / cols - Inches(0.3))
        card_height = int(available_height / rows - Inches(0.3))
        start_y = Inches(1.8)

        # 번호 색상 교대
        accent_colors = [self.color_primary, self.color_secondary, self.color_accent]

        node_idx = 0
        for r in range(rows):
            for c in range(cols):
                if node_idx >= node_count:
                    break
                node = nodes[node_idx]
                x = int(self.margin + c * (card_width + Inches(0.3)))
                y = int(start_y + r * (card_height + Inches(0.3)))

                # 카드 배경 (연한 색)
                card = slide.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    x, y, card_width, card_height,
                )
                card.fill.solid()
                card.fill.fore_color.rgb = hex_to_rgb('#F5F5F5')
                card.line.color.rgb = hex_to_rgb('#E0E0E0')

                # 번호 원
                accent = accent_colors[node_idx % len(accent_colors)]
                num_size = int(Inches(0.5))
                num_x = x + int(Inches(0.2))
                num_y = y + int(Inches(0.2))
                num_shape = slide.shapes.add_shape(
                    MSO_SHAPE.OVAL,
                    num_x, num_y, num_size, num_size,
                )
                num_shape.fill.solid()
                num_shape.fill.fore_color.rgb = hex_to_rgb(accent)
                num_shape.line.fill.background()

                tf = num_shape.text_frame
                p = tf.paragraphs[0]
                p.text = str(node_idx + 1)
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(11)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

                # 텍스트
                text_x = x + int(Inches(0.2))
                text_y = num_y + num_size + int(Inches(0.15))
                text_w = card_width - int(Inches(0.4))
                text_h = card_height - num_size - int(Inches(0.55))
                txbox = slide.shapes.add_textbox(text_x, text_y, text_w, text_h)
                tf = txbox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = self._get_node_text(node)
                p.alignment = PP_ALIGN.LEFT
                for run in p.runs:
                    run.font.size = Pt(10)
                    run.font.name = self.font_body
                    run.font.color.rgb = hex_to_rgb(self.color_text)

                node_idx += 1

    # --- 유틸리티 메서드 ---

    def _get_node_text(self, node: dict) -> str:
        """노드에서 텍스트를 가져오고 최대 길이를 적용한다."""
        text = str(node.get('text', '') or '')
        if len(text) > self.MAX_NODE_TEXT_LEN:
            text = text[:self.MAX_NODE_TEXT_LEN - 3] + '...'
        return text

    def _add_title(self, slide, title_text: str) -> None:
        """슬라이드에 제목 텍스트를 추가한다."""
        txbox = slide.shapes.add_textbox(
            int(self.margin), int(Inches(0.3)),
            int(self.slide_width - 2 * self.margin), int(Inches(0.8)),
        )
        tf = txbox.text_frame
        p = tf.paragraphs[0]
        p.text = title_text
        p.alignment = PP_ALIGN.LEFT
        for run in p.runs:
            run.font.size = self.title_size
            run.font.name = self.font_title
            run.font.bold = True
            run.font.color.rgb = hex_to_rgb(self.color_text)

    def _style_table_header_cell(self, cell) -> None:
        """표 헤더 셀 스타일을 적용한다."""
        self._set_cell_fill(cell, self.color_primary)
        p = cell.text_frame.paragraphs[0]
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size = self.body_size
            run.font.name = self.font_body

    def _set_cell_fill(self, cell, hex_color: str) -> None:
        """표 셀에 채우기 색상을 설정한다."""
        from pptx.oxml.ns import qn
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        solidFill = tcPr.makeelement(qn('a:solidFill'), {})
        srgbClr = solidFill.makeelement(
            qn('a:srgbClr'),
            {'val': hex_color.lstrip('#')},
        )
        solidFill.append(srgbClr)
        tcPr.append(solidFill)

    def _nodes_to_table_data(self, diagram_data: DiagramData) -> list[list[str]]:
        """노드 데이터를 표 형태로 변환한다."""
        nodes = diagram_data.nodes
        if not nodes:
            return []

        # 노드의 텍스트를 2열 표로 변환 (ID, 텍스트)
        table = [['항목', '내용']]
        for node in nodes:
            table.append([
                node.get('id', ''),
                node.get('text', ''),
            ])
        return table

    def _nodes_to_chart_data(self, diagram_data: DiagramData) -> dict | None:
        """노드 데이터에서 차트 데이터를 추출 시도한다."""
        nodes = diagram_data.nodes
        if not nodes:
            return None

        # 노드 텍스트에서 숫자 추출 시도
        categories = []
        values = []
        for node in nodes:
            text = node.get('text', '')
            categories.append(text)
            # 간단한 숫자 추출
            import re
            numbers = re.findall(r'[\d,]+\.?\d*', text)
            if numbers:
                val = numbers[-1].replace(',', '')
                try:
                    values.append(float(val))
                except ValueError:
                    values.append(0)
            else:
                values.append(0)

        if not any(v > 0 for v in values):
            return None

        return {
            'categories': categories,
            'series': [{'name': '값', 'values': values}],
        }

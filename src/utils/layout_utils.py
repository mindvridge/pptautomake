"""레이아웃 계산 유틸리티 - 도형 배치, 간격, 정렬 계산"""

import math
from pptx.util import Inches, Emu


# 기본 슬라이드 크기 (와이드스크린 16:9)
DEFAULT_SLIDE_WIDTH = Inches(13.333)
DEFAULT_SLIDE_HEIGHT = Inches(7.5)


def calculate_even_spacing(
    total_width: int,
    item_count: int,
    margin: int = None,
    gap: int = None,
) -> list[dict]:
    """항목을 균등 간격으로 배치한 위치 리스트를 반환한다.

    Args:
        total_width: 전체 가용 너비 (EMU)
        item_count: 배치할 항목 수
        margin: 양쪽 여백 (EMU). None이면 Inches(0.5)
        gap: 항목 간 간격 (EMU). None이면 Inches(0.3)

    Returns:
        [{'x': int, 'width': int}, ...] 리스트
    """
    if margin is None:
        margin = Inches(0.5)
    if gap is None:
        gap = Inches(0.3)

    available = total_width - 2 * margin
    total_gaps = (item_count - 1) * gap if item_count > 1 else 0
    item_width = round((available - total_gaps) / item_count)

    positions = []
    for i in range(item_count):
        x = margin + i * (item_width + gap)
        positions.append({'x': int(x), 'width': int(item_width)})
    return positions


def calculate_grid_positions(
    rows: int,
    cols: int,
    slide_width: int = None,
    slide_height: int = None,
    margin: int = None,
    gap: int = None,
    title_height: int = None,
) -> list[list[dict]]:
    """그리드 레이아웃의 셀 위치/크기를 계산한다.

    Returns:
        2D 리스트: grid[row][col] = {'x', 'y', 'width', 'height'}
    """
    if slide_width is None:
        slide_width = DEFAULT_SLIDE_WIDTH
    if slide_height is None:
        slide_height = DEFAULT_SLIDE_HEIGHT
    if margin is None:
        margin = Inches(0.5)
    if gap is None:
        gap = Inches(0.3)
    if title_height is None:
        title_height = Inches(1.2)

    avail_w = slide_width - 2 * margin
    avail_h = slide_height - title_height - margin
    cell_w = round((avail_w - (cols - 1) * gap) / cols)
    cell_h = round((avail_h - (rows - 1) * gap) / rows)

    grid = []
    for r in range(rows):
        row_positions = []
        for c in range(cols):
            x = margin + c * (cell_w + gap)
            y = title_height + r * (cell_h + gap)
            row_positions.append({
                'x': int(x), 'y': int(y),
                'width': int(cell_w), 'height': int(cell_h),
            })
        grid.append(row_positions)
    return grid


def calculate_circular_positions(
    count: int,
    center_x: int = None,
    center_y: int = None,
    radius: int = None,
) -> list[dict]:
    """원형 배치 위치를 계산한다. (순환 다이어그램용)

    Returns:
        [{'x': int, 'y': int, 'angle': float}, ...]
    """
    if center_x is None:
        center_x = DEFAULT_SLIDE_WIDTH // 2
    if center_y is None:
        center_y = DEFAULT_SLIDE_HEIGHT // 2
    if radius is None:
        radius = Inches(2.5)

    positions = []
    for i in range(count):
        angle = -math.pi / 2 + (2 * math.pi * i / count)
        x = center_x + int(radius * math.cos(angle))
        y = center_y + int(radius * math.sin(angle))
        positions.append({'x': x, 'y': y, 'angle': angle})
    return positions


def normalize_emu_to_relative(
    elements: list[dict],
    slide_width: int = None,
    slide_height: int = None,
) -> list[dict]:
    """EMU 좌표를 0~1 상대 좌표로 정규화한다."""
    if slide_width is None:
        slide_width = DEFAULT_SLIDE_WIDTH
    if slide_height is None:
        slide_height = DEFAULT_SLIDE_HEIGHT

    normalized = []
    for elem in elements:
        normalized.append({
            'x': elem['x'] / slide_width,
            'y': elem['y'] / slide_height,
            'width': elem.get('width', 0) / slide_width,
            'height': elem.get('height', 0) / slide_height,
            **{k: v for k, v in elem.items() if k not in ('x', 'y', 'width', 'height')},
        })
    return normalized

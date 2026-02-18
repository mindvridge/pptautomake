"""색상 변환 유틸리티 - HEX/RGB 변환, 색상 팔레트 분석"""

from pptx.dml.color import RGBColor

# 잘못된 HEX 입력 시 사용할 기본 색상
_FALLBACK_COLOR = '333333'


def _sanitize_hex(hex_color: str) -> str:
    """HEX 색상 문자열을 6자리로 정규화한다."""
    h = hex_color.lstrip('#')
    # 3자리 축약형 (#FFF -> FFFFFF)
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    # 6자리가 아니면 기본값
    if len(h) < 6:
        h = _FALLBACK_COLOR
    return h[:6]


def hex_to_rgb(hex_color: str) -> RGBColor:
    """HEX 색상 문자열을 python-pptx RGBColor로 변환한다.

    Args:
        hex_color: '#RRGGBB', 'RRGGBB', '#RGB' 형식의 HEX 색상
    """
    h = _sanitize_hex(hex_color)
    return RGBColor(
        int(h[0:2], 16),
        int(h[2:4], 16),
        int(h[4:6], 16),
    )


def rgb_to_hex(r: int, g: int, b: int) -> str:
    """RGB 값을 HEX 문자열로 변환한다."""
    return f"#{r:02X}{g:02X}{b:02X}"


def darken_color(hex_color: str, factor: float = 0.2) -> str:
    """색상을 어둡게 만든다. factor: 0.0(변화 없음) ~ 1.0(검정)"""
    h = _sanitize_hex(hex_color)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = int(r * (1 - factor))
    g = int(g * (1 - factor))
    b = int(b * (1 - factor))
    return rgb_to_hex(r, g, b)


def lighten_color(hex_color: str, factor: float = 0.2) -> str:
    """색상을 밝게 만든다. factor: 0.0(변화 없음) ~ 1.0(흰색)"""
    h = _sanitize_hex(hex_color)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = int(r + (255 - r) * factor)
    g = int(g + (255 - g) * factor)
    b = int(b + (255 - b) * factor)
    return rgb_to_hex(int(r), int(g), int(b))


def get_contrast_text_color(bg_hex: str) -> str:
    """배경색에 대한 대비 텍스트 색상(흰색/검정)을 반환한다."""
    h = _sanitize_hex(bg_hex)
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    return "#FFFFFF" if luminance < 0.5 else "#333333"

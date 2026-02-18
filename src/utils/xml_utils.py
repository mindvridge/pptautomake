"""PPTX XML 조작 유틸리티 - lxml 기반 XML 요소 분석"""

from lxml import etree

# PowerPoint XML 네임스페이스
NAMESPACES = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


def get_text_from_xml(xml_element) -> str:
    """XML 요소에서 모든 텍스트를 추출한다."""
    if xml_element is None:
        return ""
    texts = []
    for t_elem in xml_element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t_elem.text:
            texts.append(t_elem.text)
    return ' '.join(texts)


def get_fill_color_from_xml(xml_element) -> str | None:
    """XML 도형 요소에서 채우기 색상을 추출한다."""
    if xml_element is None:
        return None
    solid_fill = xml_element.find('.//a:solidFill/a:srgbClr', NAMESPACES)
    if solid_fill is not None:
        return f"#{solid_fill.get('val', '000000')}"
    return None


def get_shape_type_from_xml(xml_element) -> str | None:
    """XML 도형 요소에서 도형 유형(prst)을 추출한다."""
    if xml_element is None:
        return None
    prst_geom = xml_element.find('.//a:prstGeom', NAMESPACES)
    if prst_geom is not None:
        return prst_geom.get('prst')
    return None


def get_position_from_xml(xml_element) -> dict | None:
    """XML 요소에서 위치/크기 정보를 추출한다 (EMU 단위)."""
    if xml_element is None:
        return None
    xfrm = xml_element.find('.//a:xfrm', NAMESPACES)
    if xfrm is None:
        xfrm = xml_element.find('.//p:xfrm', NAMESPACES)
    if xfrm is None:
        return None

    off = xfrm.find('a:off', NAMESPACES)
    ext = xfrm.find('a:ext', NAMESPACES)
    if off is None or ext is None:
        return None

    return {
        'x': int(off.get('x', 0)),
        'y': int(off.get('y', 0)),
        'width': int(ext.get('cx', 0)),
        'height': int(ext.get('cy', 0)),
    }


def has_connector(xml_element) -> bool:
    """XML 요소 내에 커넥터(cxnSp)가 있는지 확인한다."""
    if xml_element is None:
        return False
    cxn_sp = xml_element.findall('.//p:cxnSp', NAMESPACES)
    return len(cxn_sp) > 0


def count_child_shapes(xml_element) -> int:
    """그룹 도형 내 하위 도형 수를 센다."""
    if xml_element is None:
        return 0
    sp_count = len(xml_element.findall('.//p:sp', NAMESPACES))
    grp_count = len(xml_element.findall('.//p:grpSp', NAMESPACES))
    return sp_count + grp_count


def extract_connector_relationships(xml_element) -> list[dict]:
    """커넥터 요소에서 연결 관계를 추출한다."""
    if xml_element is None:
        return []
    connections = []
    for cxn in xml_element.findall('.//p:cxnSp', NAMESPACES):
        st_cxn = cxn.find('.//a:stCxn', NAMESPACES)
        end_cxn = cxn.find('.//a:endCxn', NAMESPACES)
        if st_cxn is not None and end_cxn is not None:
            connections.append({
                'from_id': st_cxn.get('id'),
                'to_id': end_cxn.get('id'),
                'from_idx': st_cxn.get('idx'),
                'to_idx': end_cxn.get('idx'),
            })
    return connections

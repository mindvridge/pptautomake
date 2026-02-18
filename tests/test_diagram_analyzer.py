"""DiagramAnalyzer 테스트 - API 모킹 및 에러 시나리오"""

import json
import pytest
from unittest.mock import patch, MagicMock

from src.analyzer import SlideElement
from src.classifier import ClassifiedElement, ContentType, DiagramType
from src.diagram_analyzer import DiagramAnalyzer, DiagramData


@pytest.fixture
def analyzer():
    return DiagramAnalyzer({'vision_model': 'test-model', 'max_retries': 2})


def _make_image_element(image_blob=b'\x89PNG\r\n\x1a\n' + b'\x00' * 100):
    """도식 이미지 ClassifiedElement를 생성한다."""
    element = SlideElement(
        element_type='image',
        position={'x': 0, 'y': 0, 'width': 5000000, 'height': 3000000},
        content=image_blob,
        metadata={'content_type': 'image/png'},
    )
    return ClassifiedElement(
        element=element,
        content_type=ContentType.DIAGRAM_IMAGE,
        confidence=0.8,
        needs_rebuild=True,
    )


def _make_group_element():
    """그룹 도형 ClassifiedElement를 생성한다."""
    children = [
        SlideElement(
            element_type='text',
            position={'x': i * 1000000, 'y': 0, 'width': 800000, 'height': 400000},
            content=f'Step {i+1}',
        )
        for i in range(3)
    ]
    element = SlideElement(
        element_type='group_shape',
        position={'x': 0, 'y': 0, 'width': 10000000, 'height': 5000000},
        children=children,
        metadata={'child_count': 3, 'has_text': True},
    )
    return ClassifiedElement(
        element=element,
        content_type=ContentType.GROUP_DIAGRAM,
        diagram_type=DiagramType.PROCESS,
        confidence=0.7,
        needs_rebuild=True,
    )


class TestVisionAPIWithMock:
    """Vision API 모킹 테스트"""

    def test_successful_api_call(self, analyzer):
        """API 정상 호출 시 DiagramData가 반환되는지 확인"""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text=json.dumps({
            'diagram_type': 'process',
            'title': '테스트',
            'nodes': [
                {'id': 'n1', 'text': 'Step 1', 'color_fill': '#4A90D9'},
                {'id': 'n2', 'text': 'Step 2', 'color_fill': '#4A90D9'},
            ],
            'connections': [{'from': 'n1', 'to': 'n2', 'type': 'arrow'}],
            'layout': {'direction': 'horizontal'},
        }))]

        mock_client = MagicMock()
        mock_client.messages.create.return_value = mock_response
        analyzer._anthropic_client = mock_client

        result = analyzer._analyze_image_diagram(_make_image_element())

        assert result.diagram_type == 'process'
        assert len(result.nodes) == 2
        assert result.title == '테스트'

    def test_api_retry_on_failure(self, analyzer):
        """API 실패 시 재시도하는지 확인"""
        mock_client = MagicMock()
        mock_client.messages.create.side_effect = [
            Exception('Rate limit'),
            MagicMock(content=[MagicMock(text='{"diagram_type": "process", "nodes": []}')]),
        ]
        analyzer._anthropic_client = mock_client

        with patch('time.sleep'):  # sleep 건너뛰기
            result = analyzer._analyze_image_diagram(_make_image_element())

        assert mock_client.messages.create.call_count == 2
        assert result.diagram_type == 'process'

    def test_api_all_retries_fail(self, analyzer):
        """모든 재시도 실패 시 unknown 반환"""
        mock_client = MagicMock()
        mock_client.messages.create.side_effect = Exception('Server error')
        analyzer._anthropic_client = mock_client

        with patch('time.sleep'):
            result = analyzer._analyze_image_diagram(_make_image_element())

        assert mock_client.messages.create.call_count == 2  # max_retries=2
        assert result.diagram_type == 'unknown'

    def test_api_returns_invalid_json(self, analyzer):
        """API가 잘못된 JSON을 반환할 때 unknown 반환"""
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text='This is not valid JSON at all')]

        mock_client = MagicMock()
        mock_client.messages.create.return_value = mock_response
        analyzer._anthropic_client = mock_client

        result = analyzer._analyze_image_diagram(_make_image_element())

        assert result.diagram_type == 'unknown'

    def test_api_returns_json_in_code_block(self, analyzer):
        """API가 코드 블록 안에 JSON을 반환할 때 정상 파싱"""
        json_data = json.dumps({
            'diagram_type': 'flowchart',
            'nodes': [{'id': 'n1', 'text': 'Start'}],
        })
        mock_response = MagicMock()
        mock_response.content = [MagicMock(text=f'```json\n{json_data}\n```')]

        mock_client = MagicMock()
        mock_client.messages.create.return_value = mock_response
        analyzer._anthropic_client = mock_client

        result = analyzer._analyze_image_diagram(_make_image_element())

        assert result.diagram_type == 'flowchart'


class TestGroupShapeAnalysis:
    """그룹 도형 분석 테스트"""

    def test_group_shape_nodes_extracted(self, analyzer):
        """그룹 도형에서 노드가 정상 추출되는지 확인"""
        ce = _make_group_element()
        result = analyzer._analyze_group_shape_diagram(ce)

        assert result.diagram_type == 'process'
        assert len(result.nodes) == 3

    def test_group_shape_empty_children(self, analyzer):
        """하위 요소가 없는 그룹 도형 처리"""
        element = SlideElement(
            element_type='group_shape',
            position={'x': 0, 'y': 0, 'width': 10000000, 'height': 5000000},
            children=[],
        )
        ce = ClassifiedElement(
            element=element,
            content_type=ContentType.GROUP_DIAGRAM,
            diagram_type=DiagramType.PROCESS,
        )
        result = analyzer._analyze_group_shape_diagram(ce)

        assert result.nodes == []


class TestDiagramDataModel:
    """DiagramData 모델 테스트"""

    def test_default_values(self):
        data = DiagramData()
        assert data.diagram_type == 'unknown'
        assert data.node_count == 0
        assert data.table_data is None
        assert data.chart_data is None

    def test_node_count(self):
        data = DiagramData(nodes=[{'id': 'a'}, {'id': 'b'}])
        assert data.node_count == 2


class TestJsonParsing:
    """JSON 파싱 테스트"""

    def test_parse_clean_json(self, analyzer):
        result = analyzer._parse_json_response('{"key": "value"}')
        assert result == {'key': 'value'}

    def test_parse_json_with_prefix(self, analyzer):
        result = analyzer._parse_json_response('Here is the result:\n{"key": "value"}')
        assert result == {'key': 'value'}

    def test_parse_json_code_block(self, analyzer):
        text = '```json\n{"key": "value"}\n```'
        result = analyzer._parse_json_response(text)
        assert result == {'key': 'value'}

    def test_parse_invalid_json(self, analyzer):
        result = analyzer._parse_json_response('not json')
        assert result == {}

    def test_parse_empty_string(self, analyzer):
        result = analyzer._parse_json_response('')
        assert result == {}

    def test_empty_image_returns_empty_data(self, analyzer):
        """이미지 데이터가 없을 때 빈 DiagramData 반환"""
        element = SlideElement(
            element_type='image',
            position={'x': 0, 'y': 0, 'width': 100, 'height': 100},
            content=None,
            metadata={'content_type': 'image/png'},
        )
        ce = ClassifiedElement(
            element=element,
            content_type=ContentType.DIAGRAM_IMAGE,
        )
        result = analyzer.analyze(ce)
        assert result.diagram_type == 'unknown'

"""Module 3: DiagramAnalyzer - 도식 상세 분석 (Vision AI 연동)

도식으로 분류된 이미지는 Vision API로,
그룹 도형은 XML 구조 분석으로 상세 구조를 파악한다.

지원 백엔드:
  - local: 로컬 비전 모델 직접 로드 (서버/API 키 불필요, pip만으로 완결)
  - ollama: Ollama 로컬 서버 경유 (ollama 별도 설치 필요)
  - gemini: Google Gemini API (GEMINI_API_KEY 필요, 저렴)
  - anthropic: Claude Vision API (ANTHROPIC_API_KEY 필요)
"""

from __future__ import annotations

import base64
import json
import logging
import time
from dataclasses import dataclass, field
from typing import Any

from src.analyzer import SlideElement
from src.classifier import ClassifiedElement, ContentType, DiagramType
from src.utils.xml_utils import (
    get_text_from_xml,
    get_fill_color_from_xml,
    get_position_from_xml,
    extract_connector_relationships,
)

logger = logging.getLogger(__name__)

# Vision 모델에 보낼 프롬프트
VISION_ANALYSIS_PROMPT = """이 이미지는 PowerPoint 프레젠테이션의 도식/다이어그램입니다.
이 도식을 PowerPoint 네이티브 요소로 재구성하기 위해 다음을 분석해주세요:

1. 도식 유형: flowchart / process / comparison_table / hierarchy / cycle / grid / timeline / bar_chart / pie_chart / infographic
2. 구성 요소: 각 박스/노드의 텍스트, 색상(hex), 크기 비율
3. 연결 관계: 화살표/선 방향, 연결 순서
4. 레이아웃: 배치 패턴 (수평, 수직, 그리드, 방사형)
5. 스타일: 둥근 모서리 여부, 그림자, 아이콘 유무

반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 포함하지 마세요:

{
  "diagram_type": "process",
  "title": "도식 제목",
  "nodes": [
    {
      "id": "node1",
      "text": "텍스트 내용",
      "color_fill": "#4A90D9",
      "color_text": "#FFFFFF",
      "shape": "rounded_rect",
      "relative_position": {"row": 0, "col": 0},
      "relative_size": "medium",
      "icon_description": ""
    }
  ],
  "connections": [
    {"from": "node1", "to": "node2", "type": "arrow", "label": ""}
  ],
  "layout": {"direction": "horizontal", "rows": 1, "cols": 5, "spacing": "even"},
  "style": {"corner_radius": "rounded", "has_shadow": false, "border_color": "#333333"}
}"""


@dataclass
class DiagramData:
    """분석된 도식의 구조화된 데이터"""
    diagram_type: str = "unknown"
    title: str = ""
    nodes: list[dict] = field(default_factory=list)
    connections: list[dict] = field(default_factory=list)
    layout: dict = field(default_factory=dict)
    style: dict = field(default_factory=dict)
    raw_data: dict = field(default_factory=dict)
    # 표/차트용 데이터
    table_data: list[list[str]] | None = None
    chart_data: dict | None = None

    @property
    def node_count(self) -> int:
        return len(self.nodes)


class DiagramAnalyzer:
    """도식 상세 분석 모듈 - Vision API 및 XML 구조 분석"""

    # API 호출 설정
    MAX_RETRIES = 3
    RETRY_DELAYS = [2, 4, 8]  # 지수 백오프 (초)
    API_TIMEOUT = 60  # 초

    def __init__(self, config: dict | None = None):
        self.config = config or {}
        self.vision_backend = self.config.get('vision_backend', 'local')
        self.vision_model = self.config.get('vision_model', 'claude-sonnet-4-20250514')
        self.gemini_model = self.config.get('gemini_model', 'gemini-2.5-flash')
        self.ollama_model = self.config.get('ollama_model', 'llama3.2-vision')
        self.ollama_base_url = self.config.get('ollama_base_url', 'http://localhost:11434')
        self.local_model = self.config.get('local_model', 'Qwen/Qwen2.5-VL-7B-Instruct')
        self.api_timeout = self.config.get('api_timeout', self.API_TIMEOUT)
        self.max_retries = self.config.get('max_retries', self.MAX_RETRIES)
        self._anthropic_client = None
        self._gemini_client = None
        self._ollama_client = None
        self._local_model = None
        self._local_tokenizer = None

    @property
    def anthropic_client(self):
        """Anthropic API 클라이언트를 반환한다 (lazy 초기화)."""
        if self._anthropic_client is None:
            try:
                import anthropic
                self._anthropic_client = anthropic.Anthropic()
            except Exception as e:
                logger.error("Anthropic 클라이언트 초기화 실패: %s", e)
                raise
        return self._anthropic_client

    @property
    def gemini_client(self):
        """Google Gemini API 클라이언트를 반환한다 (lazy 초기화)."""
        if self._gemini_client is None:
            try:
                from google import genai
                self._gemini_client = genai.Client()
            except ImportError:
                logger.error(
                    "google-genai 패키지가 설치되지 않았습니다. "
                    "'pip install google-genai' 실행 후 다시 시도하세요."
                )
                raise
            except Exception as e:
                logger.error("Gemini 클라이언트 초기화 실패: %s", e)
                raise
        return self._gemini_client

    def _is_qwen_vl(self) -> bool:
        """현재 모델이 Qwen2.5-VL 계열인지 확인한다."""
        return 'qwen' in self.local_model.lower() and 'vl' in self.local_model.lower()

    def _load_local_model(self):
        """로컬 비전 모델을 메모리에 로드한다 (최초 1회, 이후 재사용)."""
        if self._local_model is not None:
            return

        model_id = self.local_model
        logger.info("로컬 비전 모델 로드 중: %s (최초 실행 시 다운로드)", model_id)

        try:
            import torch
        except ImportError:
            logger.error(
                "torch 패키지가 없습니다. "
                "'pip install torch torchvision' 실행 후 다시 시도하세요."
            )
            raise

        device = 'cuda' if torch.cuda.is_available() else 'cpu'
        dtype = torch.float16 if device == 'cuda' else torch.float32

        if self._is_qwen_vl():
            self._load_qwen_vl(model_id, device, dtype)
        else:
            self._load_generic_vlm(model_id, device, dtype)

        self._local_device = device
        logger.info("로컬 비전 모델 로드 완료 (device: %s)", device)

    def _load_qwen_vl(self, model_id: str, device: str, dtype):
        """Qwen2.5-VL 모델을 로드한다."""
        import torch
        try:
            from transformers import Qwen2_5_VLForConditionalGeneration, AutoProcessor
        except ImportError:
            logger.error(
                "transformers>=4.45.0이 필요합니다. "
                "'pip install -U transformers qwen-vl-utils' 실행 후 다시 시도하세요."
            )
            raise

        self._local_processor = AutoProcessor.from_pretrained(
            model_id, trust_remote_code=True,
        )
        self._local_model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
            model_id,
            trust_remote_code=True,
            torch_dtype=dtype,
            device_map="auto" if device == 'cuda' else {'': device},
        )
        self._local_tokenizer = None  # Qwen-VL은 processor 사용

    def _load_generic_vlm(self, model_id: str, device: str, dtype):
        """Moondream2 등 범용 VLM 모델을 로드한다."""
        from transformers import AutoModelForCausalLM, AutoTokenizer

        self._local_tokenizer = AutoTokenizer.from_pretrained(
            model_id, trust_remote_code=True,
        )
        self._local_model = AutoModelForCausalLM.from_pretrained(
            model_id,
            trust_remote_code=True,
            torch_dtype=dtype,
            device_map={'': device},
        )
        self._local_processor = None

    @property
    def ollama_client(self):
        """Ollama 클라이언트를 반환한다 (lazy 초기화)."""
        if self._ollama_client is None:
            try:
                import ollama
                self._ollama_client = ollama.Client(host=self.ollama_base_url)
            except ImportError:
                logger.error(
                    "ollama 패키지가 설치되지 않았습니다. "
                    "'pip install ollama' 실행 후 다시 시도하세요."
                )
                raise
            except Exception as e:
                logger.error("Ollama 클라이언트 초기화 실패: %s", e)
                raise
        return self._ollama_client

    def analyze(self, classified_element: ClassifiedElement) -> DiagramData:
        """분류된 요소를 상세 분석하여 DiagramData를 반환한다."""
        if classified_element.content_type == ContentType.DIAGRAM_IMAGE:
            return self._analyze_image_diagram(classified_element)
        elif classified_element.content_type == ContentType.GROUP_DIAGRAM:
            return self._analyze_group_shape_diagram(classified_element)
        else:
            logger.warning(
                "지원하지 않는 콘텐츠 유형: %s",
                classified_element.content_type,
            )
            return DiagramData()

    def _analyze_image_diagram(
        self, classified_element: ClassifiedElement
    ) -> DiagramData:
        """이미지 도식을 Vision API로 분석한다 (백엔드 자동 선택)."""
        element = classified_element.element
        image_blob = element.content

        if not image_blob:
            logger.warning("이미지 데이터가 없습니다")
            return DiagramData()

        content_type = element.metadata.get('content_type', 'image/png')
        # image/x-emf 등 지원하지 않는 형식 변환
        if content_type not in ('image/png', 'image/jpeg', 'image/gif', 'image/webp'):
            image_blob, content_type = self._convert_image(image_blob, content_type)

        if self.vision_backend == 'local':
            return self._call_local_vision(image_blob, content_type)
        elif self.vision_backend == 'ollama':
            return self._call_ollama_vision(image_blob, content_type)
        elif self.vision_backend == 'gemini':
            return self._call_gemini_vision(image_blob, content_type)
        else:
            return self._call_anthropic_vision(image_blob, content_type)

    def _call_anthropic_vision(
        self, image_blob: bytes, content_type: str
    ) -> DiagramData:
        """Anthropic Claude Vision API로 도식을 분석한다."""
        image_b64 = base64.b64encode(image_blob).decode('utf-8')

        for attempt in range(self.max_retries):
            try:
                response = self.anthropic_client.messages.create(
                    model=self.vision_model,
                    max_tokens=4096,
                    timeout=self.api_timeout,
                    messages=[{
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": content_type,
                                    "data": image_b64,
                                },
                            },
                            {
                                "type": "text",
                                "text": VISION_ANALYSIS_PROMPT,
                            },
                        ],
                    }],
                )

                response_text = response.content[0].text
                diagram_json = self._parse_json_response(response_text)

                if not diagram_json:
                    logger.warning(
                        "Vision API 응답 JSON 파싱 실패 (시도 %d/%d): 빈 결과",
                        attempt + 1, self.max_retries,
                    )
                    return DiagramData(diagram_type="unknown")

                return self._json_to_diagram_data(diagram_json)

            except Exception as e:
                if attempt < self.max_retries - 1:
                    delay = self.RETRY_DELAYS[min(attempt, len(self.RETRY_DELAYS) - 1)]
                    logger.warning(
                        "Anthropic API 호출 실패 (시도 %d/%d): %s - %d초 후 재시도",
                        attempt + 1, self.max_retries, e, delay,
                    )
                    time.sleep(delay)
                else:
                    logger.error(
                        "Anthropic API 분석 실패 (%d회 시도 후 포기): %s",
                        self.max_retries, e,
                    )

        return DiagramData(diagram_type="unknown")

    def _call_gemini_vision(
        self, image_blob: bytes, content_type: str
    ) -> DiagramData:
        """Google Gemini API로 도식을 분석한다 (저렴, GEMINI_API_KEY 필요)."""
        from google.genai import types

        # PIL Image로 변환
        import io
        from PIL import Image
        image = Image.open(io.BytesIO(image_blob)).convert('RGB')

        for attempt in range(self.max_retries):
            try:
                response = self.gemini_client.models.generate_content(
                    model=self.gemini_model,
                    contents=[image, VISION_ANALYSIS_PROMPT],
                    config=types.GenerateContentConfig(
                        temperature=0.1,
                        max_output_tokens=4096,
                    ),
                )

                response_text = response.text
                diagram_json = self._parse_json_response(response_text)

                if not diagram_json:
                    logger.warning(
                        "Gemini 응답 JSON 파싱 실패 (시도 %d/%d): 빈 결과",
                        attempt + 1, self.max_retries,
                    )
                    return DiagramData(diagram_type="unknown")

                return self._json_to_diagram_data(diagram_json)

            except Exception as e:
                if attempt < self.max_retries - 1:
                    delay = self.RETRY_DELAYS[min(attempt, len(self.RETRY_DELAYS) - 1)]
                    logger.warning(
                        "Gemini API 호출 실패 (시도 %d/%d): %s - %d초 후 재시도",
                        attempt + 1, self.max_retries, e, delay,
                    )
                    time.sleep(delay)
                else:
                    logger.error(
                        "Gemini API 분석 실패 (%d회 시도 후 포기): %s",
                        self.max_retries, e,
                    )

        return DiagramData(diagram_type="unknown")

    def _call_ollama_vision(
        self, image_blob: bytes, content_type: str
    ) -> DiagramData:
        """Ollama 로컬 비전 모델로 도식을 분석한다 (API 키 불필요)."""
        image_b64 = base64.b64encode(image_blob).decode('utf-8')

        for attempt in range(self.max_retries):
            try:
                response = self.ollama_client.chat(
                    model=self.ollama_model,
                    messages=[{
                        "role": "user",
                        "content": VISION_ANALYSIS_PROMPT,
                        "images": [image_b64],
                    }],
                    options={
                        "temperature": 0.1,
                        "num_predict": 4096,
                    },
                )

                response_text = response['message']['content']
                diagram_json = self._parse_json_response(response_text)

                if not diagram_json:
                    logger.warning(
                        "Ollama 응답 JSON 파싱 실패 (시도 %d/%d): 빈 결과",
                        attempt + 1, self.max_retries,
                    )
                    return DiagramData(diagram_type="unknown")

                return self._json_to_diagram_data(diagram_json)

            except Exception as e:
                if attempt < self.max_retries - 1:
                    delay = self.RETRY_DELAYS[min(attempt, len(self.RETRY_DELAYS) - 1)]
                    logger.warning(
                        "Ollama API 호출 실패 (시도 %d/%d): %s - %d초 후 재시도",
                        attempt + 1, self.max_retries, e, delay,
                    )
                    time.sleep(delay)
                else:
                    logger.error(
                        "Ollama API 분석 실패 (%d회 시도 후 포기): %s",
                        self.max_retries, e,
                    )

        return DiagramData(diagram_type="unknown")

    def _call_local_vision(
        self, image_blob: bytes, content_type: str
    ) -> DiagramData:
        """로컬 비전 모델(transformers)로 도식을 분석한다 (서버/API 키 불필요)."""
        import io
        from PIL import Image

        self._load_local_model()

        image = Image.open(io.BytesIO(image_blob)).convert('RGB')

        for attempt in range(self.max_retries):
            try:
                if self._is_qwen_vl():
                    response_text = self._generate_with_qwen_vl(image)
                elif hasattr(self._local_model, 'answer_question'):
                    # moondream2 API: model.answer_question()
                    enc_image = self._local_model.encode_image(image)
                    response_text = self._local_model.answer_question(
                        enc_image, VISION_ANALYSIS_PROMPT, self._local_tokenizer,
                    )
                else:
                    # 일반 VLM fallback (generate 기반)
                    response_text = self._generate_with_vlm(image)

                diagram_json = self._parse_json_response(response_text)

                if not diagram_json:
                    logger.warning(
                        "로컬 모델 응답 JSON 파싱 실패 (시도 %d/%d): 빈 결과",
                        attempt + 1, self.max_retries,
                    )
                    return DiagramData(diagram_type="unknown")

                return self._json_to_diagram_data(diagram_json)

            except Exception as e:
                if attempt < self.max_retries - 1:
                    delay = self.RETRY_DELAYS[min(attempt, len(self.RETRY_DELAYS) - 1)]
                    logger.warning(
                        "로컬 모델 추론 실패 (시도 %d/%d): %s - %d초 후 재시도",
                        attempt + 1, self.max_retries, e, delay,
                    )
                    time.sleep(delay)
                else:
                    logger.error(
                        "로컬 모델 분석 실패 (%d회 시도 후 포기): %s",
                        self.max_retries, e,
                    )

        return DiagramData(diagram_type="unknown")

    def _generate_with_qwen_vl(self, image) -> str:
        """Qwen2.5-VL 모델로 도식을 분석한다."""
        import torch

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "image", "image": image},
                    {"type": "text", "text": VISION_ANALYSIS_PROMPT},
                ],
            }
        ]

        text_input = self._local_processor.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True,
        )
        inputs = self._local_processor(
            text=[text_input],
            images=[image],
            padding=True,
            return_tensors="pt",
        ).to(self._local_model.device)

        with torch.no_grad():
            output_ids = self._local_model.generate(
                **inputs,
                max_new_tokens=4096,
                temperature=0.1,
                do_sample=True,
            )

        # 입력 토큰 이후만 디코딩
        input_len = inputs['input_ids'].shape[1]
        return self._local_processor.decode(
            output_ids[0][input_len:], skip_special_tokens=True,
        )

    def _generate_with_vlm(self, image) -> str:
        """일반 VLM 모델의 generate 방식으로 추론한다 (fallback)."""
        from transformers import AutoProcessor
        import torch

        model_id = self.local_model
        processor = AutoProcessor.from_pretrained(model_id, trust_remote_code=True)

        inputs = processor(
            text=VISION_ANALYSIS_PROMPT,
            images=image,
            return_tensors="pt",
        ).to(self._local_device)

        with torch.no_grad():
            output_ids = self._local_model.generate(
                **inputs,
                max_new_tokens=4096,
                temperature=0.1,
                do_sample=True,
            )

        # 입력 토큰 이후만 디코딩
        input_len = inputs['input_ids'].shape[1]
        return processor.decode(output_ids[0][input_len:], skip_special_tokens=True)

    def _convert_image(self, image_blob: bytes, content_type: str) -> tuple[bytes, str]:
        """지원하지 않는 이미지 형식을 PNG로 변환한다."""
        try:
            import io
            from PIL import Image
            img = Image.open(io.BytesIO(image_blob))
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return buf.getvalue(), 'image/png'
        except Exception as e:
            logger.warning("이미지 변환 실패 (%s): %s", content_type, e)
            return image_blob, content_type

    def _parse_json_response(self, text: str) -> dict:
        """API 응답에서 JSON을 추출한다."""
        # JSON 블록 추출 시도
        text = text.strip()
        if text.startswith('```'):
            lines = text.split('\n')
            json_lines = []
            in_block = False
            for line in lines:
                if line.startswith('```') and not in_block:
                    in_block = True
                    continue
                elif line.startswith('```') and in_block:
                    break
                elif in_block:
                    json_lines.append(line)
            text = '\n'.join(json_lines)

        # { 로 시작하는 JSON 찾기
        start = text.find('{')
        end = text.rfind('}')
        if start != -1 and end != -1:
            text = text[start:end + 1]

        try:
            return json.loads(text)
        except json.JSONDecodeError as e:
            logger.error("JSON 파싱 실패: %s", e)
            return {}

    def _json_to_diagram_data(self, data: dict) -> DiagramData:
        """JSON 딕셔너리를 DiagramData 객체로 변환한다."""
        if not data:
            return DiagramData()

        diagram = DiagramData(
            diagram_type=data.get('diagram_type', 'unknown'),
            title=data.get('title', ''),
            nodes=data.get('nodes', []),
            connections=data.get('connections', []),
            layout=data.get('layout', {}),
            style=data.get('style', {}),
            raw_data=data,
        )

        # 표 데이터가 포함된 경우
        if 'table_data' in data:
            diagram.table_data = data['table_data']
        # 차트 데이터가 포함된 경우
        if 'chart_data' in data:
            diagram.chart_data = data['chart_data']

        return diagram

    def _analyze_group_shape_diagram(
        self, classified_element: ClassifiedElement
    ) -> DiagramData:
        """그룹 도형을 XML 구조 분석으로 상세 분석한다."""
        element = classified_element.element
        children = element.children

        nodes = []
        connections = []

        # 1) 하위 도형에서 노드와 텍스트 추출
        for i, child in enumerate(children):
            text = ''
            color = None

            if child.content and isinstance(child.content, str):
                text = child.content
            elif child.xml_element is not None:
                text = get_text_from_xml(child.xml_element)
                color = get_fill_color_from_xml(child.xml_element)

            if text.strip() or color:
                node = {
                    'id': f'node{i}',
                    'text': text.strip(),
                    'color_fill': color or '#4A90D9',
                    'color_text': '#FFFFFF',
                    'shape': 'rounded_rect',
                    'relative_position': {
                        'row': 0,
                        'col': i,
                    },
                    'relative_size': 'medium',
                }
                nodes.append(node)

        # 2) 커넥터 관계 추출
        if element.xml_element is not None:
            xml_connections = extract_connector_relationships(element.xml_element)
            for conn in xml_connections:
                connections.append({
                    'from': conn.get('from_id', ''),
                    'to': conn.get('to_id', ''),
                    'type': 'arrow',
                    'label': '',
                })

        # 3) 레이아웃 추론
        direction = 'horizontal'
        if nodes:
            positions = [child.position for child in children if child.position]
            if positions:
                y_range = max(p.get('y', 0) for p in positions) - min(p.get('y', 0) for p in positions)
                x_range = max(p.get('x', 0) for p in positions) - min(p.get('x', 0) for p in positions)
                if y_range > x_range:
                    direction = 'vertical'

        diagram_type = classified_element.diagram_type.value if classified_element.diagram_type else 'process'

        return DiagramData(
            diagram_type=diagram_type,
            title='',
            nodes=nodes,
            connections=connections,
            layout={
                'direction': direction,
                'rows': 1 if direction == 'horizontal' else len(nodes),
                'cols': len(nodes) if direction == 'horizontal' else 1,
                'spacing': 'even',
            },
            style={
                'corner_radius': 'rounded',
                'has_shadow': False,
                'border_color': '#333333',
            },
        )

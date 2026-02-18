"""PowerPoint 플러그인 원클릭 설치 스크립트

pywin32 COM을 사용하여:
1. .ppam 애드인 파일 생성 (VBA + 커스텀 리본)
2. PowerPoint AddIns 폴더에 자동 복사
3. PowerPoint에 애드인 등록

실행: python setup_plugin.py
또는: setup_plugin.bat 더블클릭
"""

from __future__ import annotations

import os
import sys
import shutil
import subprocess
import zipfile
import tempfile
from pathlib import Path

# 프로젝트 루트 디렉토리
PROJECT_ROOT = Path(__file__).parent.resolve()
ADDIN_DIR = PROJECT_ROOT / 'addin'
VBA_FILE = ADDIN_DIR / 'PPTAutoMake.bas'
RIBBON_XML = ADDIN_DIR / 'customUI14.xml'


def print_header():
    """설치 안내 헤더를 출력한다."""
    print()
    print('=' * 56)
    print('  PPT AutoMake - 플러그인 설치')
    print('  PowerPoint에 "도식 변환" 탭 자동 추가')
    print('=' * 56)
    print()


def check_platform():
    """Windows 플랫폼인지 확인한다."""
    if sys.platform != 'win32':
        print('[오류] 이 설치 스크립트는 Windows에서만 실행할 수 있습니다.')
        print('       PowerPoint COM 연동이 필요합니다.')
        return False
    return True


def check_python_version():
    """Python 버전을 확인한다."""
    if sys.version_info < (3, 10):
        print(f'[오류] Python 3.10 이상이 필요합니다. (현재: {sys.version})')
        return False
    print(f'[OK] Python {sys.version_info.major}.{sys.version_info.minor}')
    return True


def install_dependencies():
    """필수 패키지를 설치한다."""
    print('[1/5] 필수 패키지 설치 중...')

    req_file = PROJECT_ROOT / 'requirements.txt'
    if not req_file.exists():
        print('  [경고] requirements.txt 파일을 찾을 수 없습니다.')
        return False

    try:
        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', '-r', str(req_file), '--quiet'],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print('  [경고] 일부 패키지 설치에 실패했습니다:')
            for line in result.stderr.strip().splitlines()[-3:]:
                print(f'         {line}')
            print(f'         수동 설치: pip install -r {req_file}')
    except Exception:
        print('  [경고] 패키지 설치 실행에 실패했습니다.')
        print(f'         pip install -r {req_file}')

    # pywin32는 반드시 설치
    try:
        import win32com.client
        print('  [OK] pywin32 설치됨')
    except ImportError:
        print('  pywin32 설치 중...')
        try:
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', 'pywin32', '--quiet'],
            )
            print('  [OK] pywin32 설치 완료')
        except subprocess.CalledProcessError:
            print('  [오류] pywin32 설치 실패. 수동으로 설치해주세요:')
            print('         pip install pywin32')
            return False

    print('  [OK] 의존성 설치 완료')
    return True


def set_environment_variable():
    """프로젝트 경로를 환경변수에 등록한다."""
    print('[2/5] 프로젝트 경로 등록 중...')

    path_str = str(PROJECT_ROOT)
    os.environ['PPTAUTOMAKE_PATH'] = path_str

    # Windows 사용자 환경변수에 영구 등록
    try:
        subprocess.check_call(
            ['setx', 'PPTAUTOMAKE_PATH', path_str],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print(f'  [OK] PPTAUTOMAKE_PATH = {path_str}')
    except (subprocess.CalledProcessError, FileNotFoundError):
        print(f'  [경고] 환경변수 자동 등록 실패')
        print(f'         수동 등록: PPTAUTOMAKE_PATH = {path_str}')

    return True


def read_vba_source() -> str:
    """VBA 소스 코드를 읽는다."""
    if not VBA_FILE.exists():
        print(f'  [오류] VBA 파일을 찾을 수 없습니다: {VBA_FILE}')
        return ''
    return VBA_FILE.read_text(encoding='utf-8')


def get_addins_folder() -> Path:
    """PowerPoint AddIns 폴더 경로를 반환한다."""
    appdata = os.environ.get('APPDATA', '')
    if appdata:
        addins_path = Path(appdata) / 'Microsoft' / 'AddIns'
    else:
        addins_path = Path.home() / 'AppData' / 'Roaming' / 'Microsoft' / 'AddIns'
    addins_path.mkdir(parents=True, exist_ok=True)
    return addins_path


def create_ppam_addin() -> str | None:
    """COM을 사용하여 .ppam 애드인 파일을 생성한다."""
    print('[3/5] 애드인 파일 생성 중...')

    import win32com.client
    import pythoncom

    ppam_path = str(get_addins_folder() / 'PPTAutoMake.ppam')
    pptm_temp = str(Path(tempfile.gettempdir()) / 'PPTAutoMake_temp.pptm')

    ppt = None
    prs = None

    try:
        # PowerPoint 인스턴스 생성 (백그라운드)
        ppt = win32com.client.Dispatch('PowerPoint.Application')

        # 빈 프레젠테이션 생성
        prs = ppt.Presentations.Add(WithWindow=False)

        # VBA 모듈 추가
        vba_code = read_vba_source()
        if not vba_code:
            return None

        # VBA 프로젝트 접근 허용 필요
        try:
            vba_project = prs.VBProject
            vba_module = vba_project.VBComponents.Add(1)  # vbext_ct_StdModule
            vba_module.Name = 'PPTAutoMake'

            # Attribute 라인 제거 (COM에서 추가 시 불필요)
            clean_code = '\n'.join(
                line for line in vba_code.splitlines()
                if not line.startswith('Attribute VB_Name')
            )
            vba_module.CodeModule.AddFromString(clean_code)
            print('  [OK] VBA 매크로 삽입 완료')
        except Exception as e:
            print(f'  [경고] VBA 자동 삽입 실패: {e}')
            print('         PowerPoint 보안 설정을 확인하세요:')
            print('         파일 → 옵션 → 보안 센터 → 보안 센터 설정')
            print('         → 매크로 설정 → "VBA 프로젝트 개체 모델에 안전하게 액세스" 체크')
            prs.Close()
            prs = None
            return _create_ppam_manual_fallback()

        # .pptm으로 임시 저장
        # ppSaveAsOpenXMLPresentationMacroEnabled = 25
        prs.SaveAs(pptm_temp, 25)
        prs.Close()
        prs = None

        # .pptm에 커스텀 리본 XML 삽입
        _inject_ribbon_xml(pptm_temp, ppam_path)

        # 임시 파일 정리
        try:
            os.remove(pptm_temp)
        except OSError:
            pass

        print(f'  [OK] 애드인 생성 완료: {ppam_path}')
        return ppam_path

    except Exception as e:
        print(f'  [오류] 애드인 생성 실패: {e}')
        if prs:
            try:
                prs.Close()
            except Exception:
                pass
        return _create_ppam_manual_fallback()

    finally:
        if ppt:
            try:
                # 열려있던 PPT가 아니라 새로 생성한 경우에만 종료
                if ppt.Presentations.Count == 0:
                    ppt.Quit()
            except Exception:
                pass


def _inject_ribbon_xml(pptm_path: str, ppam_path: str):
    """PPTM 파일에 커스텀 리본 XML을 삽입하고 PPAM으로 저장한다."""
    ribbon_content = ''
    if RIBBON_XML.exists():
        ribbon_content = RIBBON_XML.read_text(encoding='utf-8')
    else:
        ribbon_content = _get_default_ribbon_xml()

    content_types_addition = (
        '<Override PartName="/customUI/customUI14.xml" '
        'ContentType="application/xml"/>'
    )

    rels_addition = (
        '<Relationship Id="rCustomUI" '
        'Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" '
        'Target="customUI/customUI14.xml"/>'
    )

    with zipfile.ZipFile(pptm_path, 'r') as zin:
        with zipfile.ZipFile(ppam_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)

                if item == '[Content_Types].xml':
                    # 커스텀UI 콘텐츠 타입 추가
                    text = data.decode('utf-8')
                    if 'customUI14' not in text:
                        text = text.replace(
                            '</Types>',
                            f'{content_types_addition}\n</Types>',
                        )
                    zout.writestr(item, text)

                elif item == '_rels/.rels':
                    # 리본 관계 추가
                    text = data.decode('utf-8')
                    if 'rCustomUI' not in text:
                        text = text.replace(
                            '</Relationships>',
                            f'{rels_addition}\n</Relationships>',
                        )
                    zout.writestr(item, text)

                else:
                    zout.writestr(item, data)

            # 커스텀 리본 XML 추가
            zout.writestr('customUI/customUI14.xml', ribbon_content)


def _get_default_ribbon_xml() -> str:
    """기본 리본 XML을 반환한다."""
    return '''<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="pptAutoMakeTab" label="도식 변환">
        <group id="mainGroup" label="PPT AutoMake">
          <button id="btnRun"
                  label="전체 변환"
                  screentip="이미지 도식을 네이티브 요소로 변환"
                  supertip="현재 프레젠테이션의 모든 이미지 도식을 분석하고 네이티브 PowerPoint 요소(도형, 표, 차트)로 자동 변환합니다."
                  onAction="Ribbon_Run"
                  imageMso="AutomationProgramming"
                  size="large"/>
          <button id="btnAnalyze"
                  label="분석만"
                  screentip="도식 분석 결과만 확인"
                  supertip="변환 없이 도식 분석 결과만 확인합니다."
                  onAction="Ribbon_Analyze"
                  imageMso="ReviewTrackChanges"
                  size="large"/>
          <separator id="sep1"/>
          <button id="btnSelected"
                  label="선택 슬라이드"
                  screentip="특정 슬라이드만 처리"
                  supertip="슬라이드 번호를 입력하여 특정 슬라이드만 처리합니다."
                  onAction="Ribbon_SelectedSlides"
                  imageMso="SlideShowCustom"
                  size="normal"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>'''


def _create_ppam_manual_fallback() -> str | None:
    """VBA 자동 삽입이 실패한 경우 수동 설치 안내를 제공한다."""
    print()
    print('  [대안] VBA 수동 설치 방법:')
    print('  1. PowerPoint 열기')
    print('  2. Alt+F11 (VBA 편집기)')
    print('  3. 삽입 → 모듈')
    print(f'  4. {VBA_FILE} 내용 붙여넣기')
    print('  5. 파일 → 다른 이름으로 저장 → .ppam 형식')
    print(f'     저장 위치: {get_addins_folder()}')
    print()
    return None


def register_addin(ppam_path: str):
    """PowerPoint에 애드인을 등록하고 활성화한다."""
    print('[4/5] 애드인 등록 중...')

    import win32com.client

    try:
        ppt = win32com.client.Dispatch('PowerPoint.Application')

        # 기존 등록된 PPTAutoMake 애드인 제거
        for i in range(ppt.AddIns.Count, 0, -1):
            try:
                addin = ppt.AddIns.Item(i)
                if 'PPTAutoMake' in addin.Name:
                    addin.Registered = False
            except Exception:
                pass

        # 새 애드인 등록
        addin = ppt.AddIns.Add(ppam_path)
        addin.Registered = True
        addin.Loaded = True

        print(f'  [OK] 애드인 등록 완료: {ppam_path}')
        print('  [OK] PowerPoint 리본 메뉴에 "도식 변환" 탭 추가됨')
        return True

    except Exception as e:
        print(f'  [경고] 자동 등록 실패: {e}')
        print('         수동 등록 방법:')
        print('         1. PowerPoint → 파일 → 옵션 → 추가 기능')
        print('         2. 관리: "PowerPoint 추가 기능" 선택 → 이동')
        print('         3. "새로 추가" → PPTAutoMake.ppam 선택')
        return False


def check_vision_backend():
    """비전 백엔드 설정을 확인한다."""
    print('[5/5] 비전 백엔드 확인...')

    # config.yaml에서 백엔드 확인
    config_path = PROJECT_ROOT / 'config.yaml'
    backend = 'local'
    if config_path.exists():
        try:
            import yaml
            with open(config_path, 'r', encoding='utf-8') as f:
                config = yaml.safe_load(f) or {}
            backend = config.get('analysis', {}).get('vision_backend', 'local')
        except Exception:
            pass

    if backend == 'local':
        print('  [OK] 비전 백엔드: local (서버/API 키 불필요)')
        print('       최초 실행 시 비전 모델을 자동 다운로드합니다.')
        try:
            import transformers
            print(f'  [OK] transformers {transformers.__version__} 설치됨')
        except ImportError:
            print('  [경고] transformers 미설치. 실행 시 자동 설치됩니다.')
    elif backend == 'ollama':
        print('  [OK] 비전 백엔드: ollama')
        print('       Ollama가 실행 중이어야 합니다: https://ollama.com')
    elif backend == 'gemini':
        key = os.environ.get('GEMINI_API_KEY', '') or os.environ.get('GOOGLE_API_KEY', '')
        if key:
            masked = key[:8] + '...' + key[-4:]
            print(f'  [OK] 비전 백엔드: gemini (API 키: {masked})')
        else:
            print('  [경고] GEMINI_API_KEY가 설정되지 않았습니다.')
            print('         설정 방법: setx GEMINI_API_KEY "AIza..."')
            print('         https://aistudio.google.com/apikey 에서 발급')
    elif backend == 'anthropic':
        key = os.environ.get('ANTHROPIC_API_KEY', '')
        if key:
            masked = key[:8] + '...' + key[-4:]
            print(f'  [OK] 비전 백엔드: anthropic (API 키: {masked})')
        else:
            print('  [경고] ANTHROPIC_API_KEY가 설정되지 않았습니다.')
            print('         설정 방법: setx ANTHROPIC_API_KEY "sk-ant-..."')
            print('         또는 config.yaml에서 vision_backend를 local로 변경')

    return True


def print_success(ppam_path: str | None):
    """설치 완료 안내를 출력한다."""
    print()
    print('=' * 56)
    print('  설치 완료!')
    print('=' * 56)
    print()

    if ppam_path:
        print('  PowerPoint를 재시작하면 리본 메뉴에')
        print('  "도식 변환" 탭이 나타납니다.')
        print()
        print('  리본 버튼:')
        print('    [전체 변환]     - 모든 슬라이드 처리')
        print('    [분석만]        - 분석 결과만 확인')
        print('    [선택 슬라이드] - 특정 슬라이드만 처리')
    else:
        print('  배치 파일로 실행할 수 있습니다:')
        print(f'    {PROJECT_ROOT / "run_in_ppt.bat"}')

    print()
    print('  또는 명령줄에서:')
    print('    python -m src.ppt_plugin')
    print('    python -m src.ppt_plugin --slides 1,3,5')
    print('    python -m src.ppt_plugin --analyze')
    print()


def main():
    print_header()

    # 플랫폼 확인
    if not check_platform():
        sys.exit(1)

    # Python 버전 확인
    if not check_python_version():
        sys.exit(1)

    print()

    # 1. 의존성 설치
    install_dependencies()
    print()

    # 2. 환경변수 등록
    set_environment_variable()
    print()

    # 3. .ppam 애드인 생성
    ppam_path = create_ppam_addin()
    print()

    # 4. 애드인 등록
    if ppam_path:
        register_addin(ppam_path)
    else:
        print('[4/5] 애드인 등록 건너뜀 (수동 설치 필요)')
    print()

    # 5. 비전 백엔드 확인
    check_vision_backend()

    # 완료
    print_success(ppam_path)


if __name__ == '__main__':
    main()

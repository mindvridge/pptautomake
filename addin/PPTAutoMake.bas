Attribute VB_Name = "PPTAutoMake"
'===============================================================
' PPT AutoMake - PowerPoint VBA 매크로
'
' 이 매크로를 PowerPoint에 추가하면 리본 메뉴에
' "도식 변환" 버튼이 생성됩니다.
'
' 설치 방법:
'   1. PowerPoint에서 Alt+F11 (VBA 편집기)
'   2. 삽입 -> 모듈
'   3. 이 파일의 내용을 붙여넣기
'   4. 저장 후 VBA 편집기 닫기
'   5. 빠른 실행 도구 모음에 매크로 추가 (선택)
'
' 또는 install_macro.bat 실행
'===============================================================

Option Explicit

' Python 실행 경로 (환경에 맞게 수정)
Private Const PYTHON_EXE As String = "python"

'---------------------------------------------------------------
' 전체 파이프라인 실행 (분석 + 변환 + 삽입)
'---------------------------------------------------------------
Public Sub PPTAutoMake_Run()
    Dim projectPath As String
    projectPath = GetProjectPath()

    If projectPath = "" Then
        MsgBox "PPT AutoMake 프로젝트 경로를 찾을 수 없습니다." & vbCrLf & _
               "install_macro.bat를 먼저 실행해주세요.", _
               vbExclamation, "PPT 도식 자동변환"
        Exit Sub
    End If

    Dim msg As String
    msg = "현재 프레젠테이션의 이미지 도식을 분석하고" & vbCrLf & _
          "네이티브 요소로 변환합니다." & vbCrLf & vbCrLf & _
          "계속하시겠습니까?"

    If MsgBox(msg, vbYesNo + vbQuestion, "PPT 도식 자동변환") = vbNo Then
        Exit Sub
    End If

    RunPlugin projectPath, ""
End Sub

'---------------------------------------------------------------
' 분석만 실행 (변환 없이)
'---------------------------------------------------------------
Public Sub PPTAutoMake_Analyze()
    Dim projectPath As String
    projectPath = GetProjectPath()

    If projectPath = "" Then
        MsgBox "PPT AutoMake 프로젝트 경로를 찾을 수 없습니다.", _
               vbExclamation, "PPT 도식 자동변환"
        Exit Sub
    End If

    RunPlugin projectPath, "--analyze"
End Sub

'---------------------------------------------------------------
' 선택한 슬라이드만 처리
'---------------------------------------------------------------
Public Sub PPTAutoMake_SelectedSlides()
    Dim projectPath As String
    projectPath = GetProjectPath()

    If projectPath = "" Then
        MsgBox "PPT AutoMake 프로젝트 경로를 찾을 수 없습니다.", _
               vbExclamation, "PPT 도식 자동변환"
        Exit Sub
    End If

    Dim slidesStr As String
    slidesStr = InputBox("처리할 슬라이드 번호를 입력하세요:" & vbCrLf & _
                         "(예: 1,3,5-8)", _
                         "PPT 도식 자동변환 - 슬라이드 선택", "")

    If slidesStr = "" Then Exit Sub

    RunPlugin projectPath, "--slides " & slidesStr
End Sub

'---------------------------------------------------------------
' 플러그인 실행 (내부 함수)
'---------------------------------------------------------------
Private Sub RunPlugin(projectPath As String, extraArgs As String)
    Dim cmd As String
    cmd = PYTHON_EXE & " -m src.ppt_plugin " & extraArgs

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.CurrentDirectory = projectPath

    ' 명령 프롬프트 창에서 실행 (결과 확인 가능)
    wsh.Run "cmd /k """ & cmd & """", 1, False

    Set wsh = Nothing
End Sub

'---------------------------------------------------------------
' 프로젝트 경로 가져오기
'---------------------------------------------------------------
Private Function GetProjectPath() As String
    ' 환경변수에서 경로 확인
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    On Error Resume Next
    Dim envPath As String
    envPath = wsh.Environment("User")("PPTAUTOMAKE_PATH")
    On Error GoTo 0

    If envPath <> "" And Dir(envPath & "\src\ppt_plugin.py") <> "" Then
        GetProjectPath = envPath
        Set wsh = Nothing
        Exit Function
    End If

    ' 기본 경로 확인
    Dim defaultPaths As Variant
    defaultPaths = Array( _
        wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\pptautomake", _
        wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\pptautomake", _
        wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents\pptautomake", _
        "C:\pptautomake" _
    )

    Dim i As Long
    For i = LBound(defaultPaths) To UBound(defaultPaths)
        If Dir(defaultPaths(i) & "\src\ppt_plugin.py") <> "" Then
            GetProjectPath = defaultPaths(i)
            Set wsh = Nothing
            Exit Function
        End If
    Next i

    Set wsh = Nothing
    GetProjectPath = ""
End Function

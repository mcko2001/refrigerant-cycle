Attribute VB_Name = "ScrollWrapToCATIA"
Option Explicit

'===================================================================
' 대수나선 스크롤 랩 프로파일 -> CATIA V5 스케치 생성 매크로
'
' 전제조건:
'   1) 이 매크로는 scroll_wrap_profile.xlsm(매크로 사용 통합문서) 안에서
'      실행해야 합니다. (.xlsx는 매크로 저장 불가 -> 다른 이름으로 저장 시
'      "Excel 매크로 사용 통합 문서 (*.xlsm)"로 저장하세요.)
'   2) CATIA V5가 설치되어 있어야 하며, 실행 중이 아니면 매크로가 자동 실행합니다.
'   3) FixedScroll / OrbitingScroll 시트의 좌표(K,L=외곽, M,N=내곽) 구조가
'      바뀌지 않았다고 가정합니다. 시트 구조를 바꿨다면 아래 firstRow/열 문자를
'      맞춰 수정하세요.
'
' 동작:
'   현재 Inputs 시트에 입력된 값(오프셋, 크랭크각 등)을 기준으로 계산된
'   FixedScroll/OrbitingScroll 좌표를 그대로 읽어 CATIA 새 Part에
'   FS_Outer / FS_Inner / OS_Outer / OS_Inner 4개의 스케치(스플라인)로 생성합니다.
'   좌표 단위는 mm이며, CATIA Factory2D는 mm 단위를 그대로 받으므로 별도 환산은
'   필요 없습니다.
'===================================================================

Sub CreateScrollWrapInCATIA()

    Dim catApp As Object
    Dim partDoc As Object
    Dim part As Object
    Dim bodies As Object
    Dim body As Object
    Dim originElements As Object
    Dim xyPlane As Object

    On Error GoTo ErrHandler

    ' --- 1) CATIA 연결 (이미 실행 중이면 재사용, 아니면 새로 실행) ---
    On Error Resume Next
    Set catApp = GetObject(, "CATIA.Application")
    On Error GoTo ErrHandler
    If catApp Is Nothing Then
        Set catApp = CreateObject("CATIA.Application")
    End If
    catApp.Visible = True

    ' --- 2) 새 Part 문서 생성 ---
    Set partDoc = catApp.Documents.Add("Part")
    Set part = partDoc.Part
    Set bodies = part.Bodies
    Set body = bodies.Item("PartBody")

    Set originElements = part.OriginElements
    Set xyPlane = originElements.PlaneXY

    ' --- 3) 4개 플랭크 곡선을 각각 스케치로 생성 ---
    '                     시트명              X열   Y열   스케치 이름
    DrawCurveFromSheet part, body, xyPlane, "FixedScroll", "K", "L", "FS_Outer"
    DrawCurveFromSheet part, body, xyPlane, "FixedScroll", "M", "N", "FS_Inner"
    DrawCurveFromSheet part, body, xyPlane, "OrbitingScroll", "K", "L", "OS_Outer"
    DrawCurveFromSheet part, body, xyPlane, "OrbitingScroll", "M", "N", "OS_Inner"

    ' --- 4) 갱신 및 화면 맞춤 ---
    part.Update
    On Error Resume Next
    catApp.ActiveWindow.ActiveViewer.Reframe
    On Error GoTo ErrHandler

    MsgBox "스크롤 랩 프로파일 4개 곡선(FS_Outer/FS_Inner/OS_Outer/OS_Inner)을 " & _
           "CATIA Part에 생성했습니다." & vbCrLf & _
           "(현재 Inputs 시트 크랭크각 = " & ThisWorkbook.Sheets("Inputs").Range("C14").Value & "도 기준)", _
           vbInformation, "완료"
    Exit Sub

ErrHandler:
    MsgBox "오류가 발생했습니다: " & Err.Description & vbCrLf & _
           "CATIA V5가 설치/실행 가능한 상태인지, 시트 구조가 맞는지 확인하세요.", _
           vbCritical, "CreateScrollWrapInCATIA 오류"
End Sub


'-------------------------------------------------------------------
' 지정한 시트의 X,Y 좌표 열을 읽어 CATIA 스케치(스플라인)로 생성
'-------------------------------------------------------------------
Private Sub DrawCurveFromSheet(part As Object, body As Object, xyPlane As Object, _
                                sheetName As String, xCol As String, yCol As String, _
                                sketchName As String)

    Const firstRow As Long = 4   ' 좌표 데이터 시작행 (헤더 다음 행)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, k As Long, n As Long
    Dim sketches As Object
    Dim sketch As Object
    Dim factory2D As Object
    Dim pts() As Double

    Set ws = ThisWorkbook.Sheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, xCol).End(xlUp).Row

    n = lastRow - firstRow + 1
    If n < 2 Then
        MsgBox sheetName & " 시트에서 " & xCol & "열 좌표를 찾지 못했습니다. 열 구성을 확인하세요.", vbExclamation
        Exit Sub
    End If

    ReDim pts(1 To n * 2)
    k = 1
    For i = firstRow To lastRow
        pts(k) = ws.Cells(i, xCol).Value      ' mm
        pts(k + 1) = ws.Cells(i, yCol).Value  ' mm
        k = k + 2
    Next i

    Set sketches = body.Sketches
    Set sketch = sketches.Add(xyPlane)
    sketch.Name = sketchName

    sketch.OpenEdition
    Set factory2D = sketch.Factory2D
    factory2D.CreateSpline pts
    sketch.CloseEdition

End Sub

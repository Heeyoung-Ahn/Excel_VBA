Attribute VB_Name = "sb_RemoveHtmlTags"
Option Explicit

'------------------------------
'  Html Tag를 지우는 코드
'------------------------------
Sub RemoveHtmlTags()
    Dim r As Range
    Dim selectedRng As Range
    
    On Error Resume Next
        Set selectedRng = Application.InputBox("HTML을 제거할 영역을 선택하세요.", "HTML 제거 도구", Type:=8)
    On Error GoTo 0
    If selectedRng Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(selectedRng) = 0 Then Exit Sub
    
    With CreateObject("vbscript.regexp") '정규표현식을 위한 개체 설정
        .Pattern = "\<.*?\>"
        .Global = True
        For Each r In selectedRng
            r.Value = Replace(r, "</p><p>", "" & Chr(10) & "") '줄바꿈 태그 유지
            r.Value = Replace(r, "&amp;", "&") '& 유지
            r.Value = Replace(.Replace(r.Value, ""), "-&gt;", "→") '방향키 유지
        Next r
    End With
    
End Sub

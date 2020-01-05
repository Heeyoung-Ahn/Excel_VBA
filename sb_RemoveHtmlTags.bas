Attribute VB_Name = "sb_RemoveHtmlTags"
Option Explicit

'------------------------------
'  Html Tag를 지우는 코드
'------------------------------
Sub RemoveHtmlTags()

    Dim r As Range

    With CreateObject("vbscript.regexp") '정규표현식을 위한 개체 설정
        .Pattern = "\<.*?\>"
        .Global = True
        For Each r In Selection
            r.Value = Replace(r, "</p><p>", "" & Chr(10) & "") '줄바꿈 태그 유지
            r.Value = Replace(r, "&amp;", "&") '& 유지
            r.Value = Replace(.Replace(r.Value, ""), "-&gt;", "→") '방향키 유지
        Next r
    End With
    
End Sub

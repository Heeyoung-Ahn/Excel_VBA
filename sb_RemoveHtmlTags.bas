Attribute VB_Name = "sb_RemoveHtmlTags"
Option Explicit

'------------------------------
'  Html Tag�� ����� �ڵ�
'------------------------------
Sub RemoveHtmlTags()
    Dim r As Range
    Dim selectedRng As Range
    
    On Error Resume Next
        Set selectedRng = Application.InputBox("HTML�� ������ ������ �����ϼ���.", "HTML ���� ����", Type:=8)
    On Error GoTo 0
    If selectedRng Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(selectedRng) = 0 Then Exit Sub
    
    With CreateObject("vbscript.regexp") '����ǥ������ ���� ��ü ����
        .Pattern = "\<.*?\>"
        .Global = True
        For Each r In selectedRng
            r.Value = Replace(r, "</p><p>", "" & Chr(10) & "") '�ٹٲ� �±� ����
            r.Value = Replace(r, "&amp;", "&") '& ����
            r.Value = Replace(.Replace(r.Value, ""), "-&gt;", "��") '����Ű ����
        Next r
    End With
    
End Sub

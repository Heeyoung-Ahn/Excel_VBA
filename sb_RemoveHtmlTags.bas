Attribute VB_Name = "sb_RemoveHtmlTags"
Option Explicit

'------------------------------
'  Html Tag�� ����� �ڵ�
'------------------------------
Sub RemoveHtmlTags()

    Dim r As Range

    With CreateObject("vbscript.regexp") '����ǥ������ ���� ��ü ����
        .Pattern = "\<.*?\>"
        .Global = True
        For Each r In Selection
            r.Value = Replace(r, "</p><p>", "" & Chr(10) & "") '�ٹٲ� �±� ����
            r.Value = Replace(r, "&amp;", "&") '& ����
            r.Value = Replace(.Replace(r.Value, ""), "-&gt;", "��") '����Ű ����
        Next r
    End With
    
End Sub

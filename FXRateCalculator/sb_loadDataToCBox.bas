Attribute VB_Name = "sb_loadDataToCBox"
Option Explicit

'----------------------------------------------------------------
'  �޺��ڽ� ������
'    - loadDataToCBox(�޺��ڽ�, SQL��, DB, Form)
'----------------------------------------------------------------
Sub loadDataToCBox(argCboBox As MSForms.ComboBox, argSQL As String, argDB As String, argFormNM As String)
    Dim i As Integer, j As Integer
    Dim listData() As String
    
    Call connectTaskDB
    callDBtoRS "loadDataToCBox", argDB, argSQL, argFormNM, "�޺��ڽ�������"

    If rs.EOF Then
        'MsgBox argFormNM & "�� " & argCboBox.Name & "�� ������ �ڷᰡ �����ϴ�.", vbInformation, Banner
        argCboBox.Clear
        disconnectALL
        Exit Sub
    End If
    
    ReDim listData(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB���� ��ȯ�� �迭�� ũ�� ����: ���ڵ���� ���ڵ� ��, �ʵ� ��
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        For j = 0 To rs.Fields.Count - 1
            listData(i, j) = rs.Fields(j).Value
        Next j
        rs.MoveNext
    Next i
    Call disconnectALL
    
    '//listData �迭�� ��ȯ�� Data�� �޺��ڽ��� ������
    argCboBox.List = listData
End Sub

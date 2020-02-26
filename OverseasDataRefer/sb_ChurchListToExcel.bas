Attribute VB_Name = "sb_ChurchListToExcel"
Option Explicit

'--------------------------
'  ������ȹDB ���� ��ȯ
'--------------------------
Sub ChurchListtoExcel()
    
    Dim tableNM As String, dbNM As String
    Dim strSQL As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    
    '//�α���üũ
    If checkLogin = 0 Then
        MsgBox "���� �α��� ���ּ���." & Space(10), vbInformation, banner
        Exit Sub
    End If
    
    '//db�� ����
    tableNM = "overseas.v_churches" '//db��.���̺�� - �����ڡ�
    dbNM = "overseas" '//�����ڡ�
    
    '//DB����
    connectTaskDB
    
    '//Select��-�����ڡ�
    strSQL = "SELECT * FROM " & tableNM & " WHERE `���μ�` = " & SText(user_dept) & ";"
    
    '//SQL�� �����ϰ� ��ȸ�� �ڷḦ ���ڵ�¿� ����
    callDBtoRS "gospelDBtoExcel", tableNM, strSQL, , "��ȸ����Ʈ������ȯ"
    If rs.EOF = True Then
        MsgBox "��ȸ ���ǿ� �´� �ڷᰡ �����ϴ�.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
        
    '//������ �ڷ� ��������
    Optimization
    Sheet3.Cells(1, 1).CurrentRegion.Offset(1).Delete shift:=xlUp
    Sheet3.Activate
    For i = 0 To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    '����
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).AutoFilter
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 13), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    Normal
    
    Cells(2, 1).Select
    
    '//�������, ������
    '�αױ��
    strSQL = "INSERT INTO common.logs(procedure_nm, table_nm, sql_script, error_cd, job_nm, affectedCount, user_id) " & _
                  "Values('ChurchListtoExcel', " & SText(tableNM) & ", " & SText(strSQL) & ", 0, '��ȸ����Ʈ������ȯ', " & rs.RecordCount & ", " & user_id & ");"
    executeSQL "writeLog", "common.logs", strSQL, , "�αױ��"
    disconnectALL
    '�������
    MsgBox "��ȸ����Ʈ ��ȸ�� �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub




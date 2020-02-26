Attribute VB_Name = "sb_GospelRegionDBtoExcel"
Option Explicit

'--------------------------
'  ������ȹDB ���� ��ȯ
'--------------------------
Sub GospelRegionDBtoExcel()
    
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
    tableNM = "overseas.v_gospel_regions" '//db��.���̺�� - �����ڡ�
    dbNM = "overseas" '//�����ڡ�
    
    '//DB����
    connectTaskDB
    
    '//Select��-�����ڡ�
    strSQL = "SELECT * FROM " & tableNM & " WHERE `���μ�` = " & SText(user_dept) & ";"
    
    '//SQL�� �����ϰ� ��ȸ�� �ڷḦ ���ڵ�¿� ����
    callDBtoRS "gospelDBtoExcel", tableNM, strSQL, , "������ȹ������ȯ"
    If rs.EOF = True Then
        MsgBox "��ȸ ���ǿ� �´� �ڷᰡ �����ϴ�.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
        
    '//������ �ڷ� ��������
    Optimization
    Sheet2.Cells(1, 1).CurrentRegion.Offset(1).Delete shift:=xlUp
    Sheet2.Activate
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
        .SortFields.Add Key:=Cells(1, 5), Order:=xlAscending
        .SortFields.Add Key:=Cells(1, 6), Order:=xlAscending
        .SortFields.Add Key:=Cells(1, 7), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    Normal
    
    Cells(2, 1).Select
    
    '//�������, ������
    '�αױ��-�����ڡ�
    strSQL = "INSERT INTO common.logs(procedure_nm, table_nm, sql_script, error_cd, job_nm, affectedCount, user_id) " & _
                  "Values('GospelRegionDBtoExcel', " & SText(tableNM) & ", " & SText(strSQL) & ", 0, '������ȹ�ڷῢ����ȯ', " & rs.RecordCount & ", " & user_id & ");"
    executeSQL "writeLog", "common.logs", strSQL, , "�αױ��"
    disconnectALL
    '�������
    MsgBox "������ȹ�ڷ� ��ȸ�� �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub


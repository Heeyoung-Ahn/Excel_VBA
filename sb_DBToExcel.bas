Attribute VB_Name = "sb_DBToExcel"
Option Explicit

Public Const banner As String = "Excel To Database Tool V1.0"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const IPAddress As String = "IP�ּ�" 'DB IP Address�ڡ�
Public Const DBPassword As String = "Password" '��й�ȣ�ڡ�
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset

'-----------------------------------------------------------------------------
'  DB���� Select������ �˻��� �ڷḦ ���ڵ�¿� ��Ƽ� ������ ��ȯ
'    - excel_export(���Ͽ��¿���)
'    - ��ũ���� ���� ����ȭ�鿡 ����
'    - ���Ͽ����� true�� ���� �� ����
'-----------------------------------------------------------------------------
Sub excel_export(Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSQL As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    
    '//db�� ����
    tableNM = "accounts.transactions" '//db��.���̺�� - �����ڡ�
    dbNM = "accounts" '//�����ڡ�
    
    '//DB����
    connectDB IPAddress, dbNM, "root", DBPassword
    
    '//Select��
    strSQL = "SELECT ���� FROM " & tableNM & " WHERE ���ǽ�" & ";"
    
    '//SQL�� �����ϰ� ��ȸ�� �ڷḦ ���ڵ�¿� ����
    callDBtoRS strSQL
    If rs.EOF = True Then
        MsgBox "��ȸ ���ǿ� �´� �ڷᰡ �����ϴ�.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
        
    '//������ �ڷ� ��������
    'Optimization
    Workbooks.Add
    For i = i To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    fileNM = GetDesktopPath() & tableNM & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".xlsx"
    Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs Filename:=fileNM
    Application.DisplayAlerts = True
    If FileOpen = False Then
        ActiveWorkbook.Close
    End If
    'Normal
    
    '//�������, ������
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "����ȭ�鿡 ������ �����Ǿ����ϴ�." & vbNewLine & vbNewLine & _
        " - �����̸�: " & fileSNM, vbInformation, banner
    disconnectALL
End Sub

'-----------------------------------------------
'  DB����
'    - connectDB(���� IP, ��Ű��, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'--------------------------
'  DB �� RS ���� ����
'--------------------------
Sub disconnectALL()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub

'-----------------------------------------------
'  ���ڵ�� ���� �� ������ ��ȯ
'-----------------------------------------------
Sub callDBtoRS(SQLScript As String)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
End Sub

'------------------------------------------------
'  SQL ���ϸ�Ī �˻��� ó��('%�˻���%')
'------------------------------------------------
Public Function PText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        PText = "'%%'"
    Else
        PText = "'%" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "%'"
    End If
End Function

'------------------------------------------------
'  SQL ��Į���Ī �˻��� ó��('�˻���')
'------------------------------------------------
Public Function SText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        SText = "''"
    Else
        SText = "'" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "'"
    End If
End Function

'--------------------
'  ��ũ�� ����ȭ
'--------------------
Sub Optimization()
On Error Resume Next
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
On Error GoTo 0
End Sub

'-------------------------
'  ��ũ�� ����ȭ ����
'-------------------------
Sub Normal()
On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
On Error GoTo 0
End Sub

'-------------------------
'  ����ȭ�� ��� ��ȸ
'-------------------------
Public Function GetDesktopPath(Optional BackSlash As Boolean = True)
    Dim oWSHShell As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    If BackSlash = True Then
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
    Else
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
    End If
    
    Set oWSHShell = Nothing
End Function

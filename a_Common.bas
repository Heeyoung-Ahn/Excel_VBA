Attribute VB_Name = "A_Common"
Option Explicit

Public Const Banner As String = "���α׷� ��Ī"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const programv As String = "Program Version"
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset
Public connIP As String, connDB As String, connUN As String, connPW As String '//Task DB ���� ����
Public user_id As Integer '������ڵ�
Public user_gb As String '����ڱ���(SA, AM, MG, WP)
Public user_nm As String '������̸�
Public checkLogin As Integer '�α��� ���� 0: �α��� ����, 1 = �α���
Public Const commonPW As String = "Password"
Public cuCode As Integer, pjCode As Integer

'-----------------------
'  Common DB����
'-----------------------
Sub connectCommonDB()
    connectDB "IP Address", "DB Name", "ID", commonPW
End Sub

'-------------------
'  Task DB����
'-------------------
Sub connectTaskDB()
    connectDB connIP, connDB, connUN, connPW
End Sub

'-----------------------------------------------
'  DB���� ���ν���
'    - connectDB(���� IP, ��Ű��, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'---------------------------------------------------------------------
'  ���ڵ�� ���� �� ������ ��ȯ
'    - calDBtoRS(���ν�����, ���̺��, SQL��, ���̸�, ���̸�)
'    - �����߻� �� ���� �ڵ鸵 �� �α� ���
'    - �����߻� ���ϸ� �� ���� ���ν������� �α� ���(�ʿ� ��)
'---------------------------------------------------------------------
Sub callDBtoRS(ProcedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional JobNM As String = "������ ��ȸ")
On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
    Exit Sub
    
ErrHandler:
    ErrHandle ProcedureNM, tableNM, SQLScript, formNM, JobNM
    writeLog ProcedureNM, tableNM, SQLScript, 1, formNM, JobNM '//�����ڵ� 1
End Sub

'-------------------------------------------------------------------------------------
'  SQL���� �����ϰ� ������ ������ ���� ���ڵ� ���� ��ȯ
'    - executeSQL(���ν�����, ���̺��, SQL��, ���̸�(�ɼ�), ���̸�(�ɼ�))
'    - SQL�� ���� ��� ���� ���θ� �˱� ���� ���� ���� ���ڵ� �� ����
'    - �����߻� �� ���� �ڵ鸵 �� �α� ���
'-------------------------------------------------------------------------------------
Public Function executeSQL(ProcedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional JobNM As String = "��Ÿ") As Long
On Error GoTo ErrHandler

    Dim affectedCount As Long
    
    conn.Execute CommandText:=SQLScript, recordsaffected:=affectedCount
    executeSQL = affectedCount
    Exit Function
    
ErrHandler:
    ErrHandle ProcedureNM, tableNM, SQLScript, formNM, JobNM
    writeLog ProcedureNM, tableNM, SQLScript, 1, formNM, JobNM '//�����ڵ� 1
End Function

'--------------------------
'  DB �� RS ���� ����
'--------------------------
Sub disconnectRS()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
    On Error GoTo 0
End Sub
Sub disconnectDB()
    On Error Resume Next
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub
Sub disconnectALL()
    On Error Resume Next
        rs.Close
        Set rs = Nothing
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
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

'---------------------------------------------
'  SQL ��Į���Ī �˻��� ó��('�˻���')
'---------------------------------------------
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

'----------------
'  ��üȭ��On
'----------------
Sub FullscreenOn()
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
End Sub

'----------------
'  ��üȭ��Off
'----------------
Sub FullscreenOff()
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

'---------------------
'  ����ȭ�� �����
'---------------------
Sub HideExcel()
    Application.Visible = False
End Sub

'---------------------
'  ����ȭ�� ���̱�
'---------------------
Sub ShowExcel()
    Application.Visible = True
End Sub

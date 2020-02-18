Attribute VB_Name = "sb_EXcelToDB"
Option Explicit

Public Const banner As String = "Excel To Database Tool V1.0"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const IPAddress As String = "IP�ּ�" 'DB IP Address�ڡ�
Public Const DBPassword As String = "Password" 'SA ��й�ȣ�ڡ�
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset

'-----------------------------------------------
'  DB����
'    - connectDB(���� IP, ��Ű��, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'-----------------
'  DB�������
'-----------------
Sub disconnectDB()
    On Error Resume Next
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
'  SQL���� �����ϰ� ������ ������ ���� ���ڵ� ���� ��ȯ
'------------------------------------------------------------------
Public Function excuteSQL(SQLScript As String) As Long
    Dim affectedCount As Long
    conn.Execute CommandText:=SQLScript, recordsaffected:=affectedCount
    excuteSQL = affectedCount
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

'---------------------------------------------------
'  ���� �ڷḦ DB�� Insert
'    - Table �ʱ�ȭ
'    - ���� ���ڵ� �ϳ��� Insert
'----------------------------------------------------
Sub insertDataToDB()
    
    Dim shtNM As String, tableNM As String, strSQL As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Integer, j As Long, k As Integer
    
    '//Sheet��, Table�� �Է� �ޱ�
    shtNM = "ch_accounts" '//�����ڡ�
    tableNM = "church_account.accounts" '//�����ڡ�
    
    '//DB����
    connectDB IPAddress, "common", "root", DBPassword
    
    '//Table �ʱ�ȭ
    strSQL = "TRUNCATE TABLE " & tableNM
    affectedCount = excuteSQL(strSQL)
    
    '/�迭 ũ�� ����
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert(�����ڡ�)
    '  - �ʵ� �� ��ŭ �߰��ؼ� ����
    '  - ������ NULL���� �ְ��� �� ��� "''"�� "NULL"�� ����
    '  - �⺻���� �������� "Default" & "," & _
    '  - TimeStamp�� ��������: "CURRENT_TIMESTAMP()" & ");"
    For i = 0 To cntRecord - 1
        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & ");"
        affectedCount = excuteSQL(strSQL)
        k = k + affectedCount
    Next i
    MsgBox k & "���� ���ڵ尡 '" & tableNM & "'�� �߰��Ǿ����ϴ�.", vbInformation, banner
    
    '//���� ����
    disconnectDB
End Sub


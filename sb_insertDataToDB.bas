Attribute VB_Name = "sb_insertDataToDB"
Option Explicit

Public Const banner As String = "Excel To Database Tool V1.0"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const IPAddress As String = "IP�ּ�" 'DB IP Address
Public Const DBPassword As String = "Password" 'SA ��й�ȣ
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset

'-------------------
'  Task DB����
'-------------------
Sub connectTaskDB()
    connectDB IPAddress, "common", "root", DBPassword
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

'------------------------------------------------------------------
'  SQL���� �����ϰ� ������ ������ ���� ���ڵ� ���� ��ȯ
'------------------------------------------------------------------
Public Function excuteSQL(SQLScript As String) As Long
    Dim affectedCount As Long
    conn.Execute CommandText:=SQLScript, recordsaffected:=affectedCount
    excuteSQL = affectedCount
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
    
    connectTaskDB
    
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
    MsgBox k & "���� ���ڵ尡 �߰��Ǿ����ϴ�.", vbInformation, banner
    
    disconnectALL
End Sub

'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
'                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
'                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & ");"

'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
'                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
'                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
'                       IIf((IsNull(Values(i, 7)) Or Values(i, 7) = ""), "''", (IIf(IsNumeric(Values(i, 7)), Values(i, 7), SText(Values(i, 7))))) & "," & _
'                       IIf((IsNull(Values(i, 8)) Or Values(i, 8) = ""), "''", (IIf(IsNumeric(Values(i, 8)), Values(i, 8), SText(Values(i, 8))))) & ");"

'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & ");"

'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
'                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & ");"

'//relation table �ڷ� ���ε��
'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       "Default" & "," & _
'                       "Default" & "," & _
'                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
'                       "CURRENT_TIMESTAMP()" & ");"

'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
'                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
'                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
'                       IIf((IsNull(Values(i, 7)) Or Values(i, 7) = ""), "''", (IIf(IsNumeric(Values(i, 7)), Values(i, 7), SText(Values(i, 7))))) & "," & _
'                       IIf((IsNull(Values(i, 8)) Or Values(i, 8) = ""), "''", (IIf(IsNumeric(Values(i, 8)), Values(i, 8), SText(Values(i, 8))))) & "," & _
'                       IIf((IsNull(Values(i, 9)) Or Values(i, 9) = ""), "''", (IIf(IsNumeric(Values(i, 9)), Values(i, 9), SText(Values(i, 9))))) & "," & _
'                       IIf((IsNull(Values(i, 10)) Or Values(i, 10) = ""), "''", (IIf(IsNumeric(Values(i, 10)), Values(i, 10), SText(Values(i, 10))))) & "," & _
'                       IIf((IsNull(Values(i, 11)) Or Values(i, 11) = ""), "''", (IIf(IsNumeric(Values(i, 11)), Values(i, 11), SText(Values(i, 11))))) & "," & _
'                       IIf((IsNull(Values(i, 12)) Or Values(i, 12) = ""), "''", (IIf(IsNumeric(Values(i, 12)), Values(i, 12), SText(Values(i, 12))))) & "," & _
'                       IIf((IsNull(Values(i, 13)) Or Values(i, 13) = ""), "''", (IIf(IsNumeric(Values(i, 13)), Values(i, 13), SText(Values(i, 13))))) & "," & _
'                       IIf((IsNull(Values(i, 14)) Or Values(i, 14) = ""), "''", (IIf(IsNumeric(Values(i, 14)), Values(i, 14), SText(Values(i, 14))))) & "," & _
'                       IIf((IsNull(Values(i, 15)) Or Values(i, 15) = ""), "''", (IIf(IsNumeric(Values(i, 15)), Values(i, 15), SText(Values(i, 15))))) & ");"

'//����޾��ε�
'  - ������������ ���� �ٿ�ε�
'  - 1���� �ڵ� �� �߰�
'  - ��޿� ' ���' �� ''���� ����: ����� DB �ڷ�����  INT
'  - ���ε� ����
'        strSQL = "INSERT INTO " & tableNM & " VALUES(" & _
'                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
'                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
'                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
'                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
'                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
'                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
'                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
'                       IIf((IsNull(Values(i, 7)) Or Values(i, 7) = ""), "''", (IIf(IsNumeric(Values(i, 7)), Values(i, 7), SText(Values(i, 7))))) & "," & _
'                       IIf((IsNull(Values(i, 8)) Or Values(i, 8) = ""), "''", (IIf(IsNumeric(Values(i, 8)), Values(i, 8), SText(Values(i, 8))))) & "," & _
'                       IIf((IsNull(Values(i, 9)) Or Values(i, 9) = ""), "''", (IIf(IsNumeric(Values(i, 9)), Values(i, 9), SText(Values(i, 9))))) & "," & _
'                       IIf((IsNull(Values(i, 10)) Or Values(i, 10) = ""), "''", (IIf(IsNumeric(Values(i, 10)), Values(i, 10), SText(Values(i, 10))))) & "," & _
'                       IIf((IsNull(Values(i, 11)) Or Values(i, 11) = ""), "''", (IIf(IsNumeric(Values(i, 11)), Values(i, 11), SText(Values(i, 11))))) & "," & _
'                       IIf((IsNull(Values(i, 12)) Or Values(i, 12) = ""), "''", (IIf(IsNumeric(Values(i, 12)), Values(i, 12), SText(Values(i, 12))))) & "," & _
'                       IIf((IsNull(Values(i, 13)) Or Values(i, 13) = ""), "''", (IIf(IsNumeric(Values(i, 13)), Values(i, 13), SText(Values(i, 13))))) & "," & _
'                       IIf((IsNull(Values(i, 14)) Or Values(i, 14) = ""), "''", (IIf(IsNumeric(Values(i, 14)), Values(i, 14), SText(Values(i, 14))))) & "," & _
'                       IIf((IsNull(Values(i, 15)) Or Values(i, 15) = ""), "''", (IIf(IsNumeric(Values(i, 15)), Values(i, 15), SText(Values(i, 15))))) & "," & _
'                       IIf((IsNull(Values(i, 16)) Or Values(i, 16) = ""), "''", (IIf(IsNumeric(Values(i, 16)), Values(i, 16), SText(Values(i, 16))))) & "," & _
'                       IIf((IsNull(Values(i, 17)) Or Values(i, 17) = ""), "''", (IIf(IsNumeric(Values(i, 17)), Values(i, 17), SText(Values(i, 17))))) & "," & _
'                       IIf((IsNull(Values(i, 18)) Or Values(i, 18) = ""), "''", (IIf(IsNumeric(Values(i, 18)), Values(i, 18), SText(Values(i, 18))))) & "," & _
'                       IIf((IsNull(Values(i, 19)) Or Values(i, 19) = ""), "''", (IIf(IsNumeric(Values(i, 19)), Values(i, 19), SText(Values(i, 19))))) & "," & _
'                       IIf((IsNull(Values(i, 20)) Or Values(i, 20) = ""), "''", (IIf(IsNumeric(Values(i, 20)), Values(i, 20), SText(Values(i, 20))))) & ");"

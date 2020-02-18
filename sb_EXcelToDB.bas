Attribute VB_Name = "sb_EXcelToDB"
Option Explicit

Public Const banner As String = "Excel To Database Tool V1.0"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const IPAddress As String = "IP주소" 'DB IP Address★★
Public Const DBPassword As String = "Password" 'SA 비밀번호★★
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset

'-----------------------------------------------
'  DB연결
'    - connectDB(서버 IP, 스키마, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'-----------------
'  DB연결끊기
'-----------------
Sub disconnectDB()
    On Error Resume Next
        conn.Close
        Set conn = Nothing
    On Error GoTo 0
End Sub

'------------------------------------------------------------------
'  SQL문을 실행하고 실행결과 영향을 받은 레코드 수를 반환
'------------------------------------------------------------------
Public Function excuteSQL(SQLScript As String) As Long
    Dim affectedCount As Long
    conn.Execute CommandText:=SQLScript, recordsaffected:=affectedCount
    excuteSQL = affectedCount
End Function

'------------------------------------------------
'  SQL 스칼라매칭 검색어 처리('검색어')
'------------------------------------------------
Public Function SText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        SText = "''"
    Else
        SText = "'" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "'"
    End If
End Function

'---------------------------------------------------
'  엑셀 자료를 DB에 Insert
'    - Table 초기화
'    - 엑셀 레코드 하나씩 Insert
'----------------------------------------------------
Sub insertDataToDB()
    
    Dim shtNM As String, tableNM As String, strSQL As String
    Dim affectedCount As Long
    Dim Values() As String
    Dim cntField As Integer, cntRecord As Long, i As Integer, j As Long, k As Integer
    
    '//Sheet명, Table명 입력 받기
    shtNM = "ch_accounts" '//수정★★
    tableNM = "church_account.accounts" '//수정★★
    
    '//DB연결
    connectDB IPAddress, "common", "root", DBPassword
    
    '//Table 초기화
    strSQL = "TRUNCATE TABLE " & tableNM
    affectedCount = excuteSQL(strSQL)
    
    '/배열 크기 지정
    Sheets(shtNM).Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Values(cntRecord - 1, cntField - 1)
   
    '//DB로 이동할 자료 Values 배열에 저장
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//레코드별로 Insert(수정★★)
    '  - 필드 수 만큼 추가해서 진행
    '  - 공백대신 NULL값을 넣고자 할 경우 "''"을 "NULL"로 수정
    '  - 기본값을 넣으려면 "Default" & "," & _
    '  - TimeStamp를 넣으려면: "CURRENT_TIMESTAMP()" & ");"
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
    MsgBox k & "개의 레코드가 '" & tableNM & "'에 추가되었습니다.", vbInformation, banner
    
    '//연결 끊기
    disconnectDB
End Sub


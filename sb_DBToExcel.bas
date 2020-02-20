Attribute VB_Name = "sb_DBToExcel"
Option Explicit

Public Const banner As String = "Excel To Database Tool V1.0"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const IPAddress As String = "IP주소" 'DB IP Address★★
Public Const DBPassword As String = "Password" '비밀번호★★
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset

'-----------------------------------------------------------------------------
'  DB에서 Select문으로 검색한 자료를 레코드셋에 담아서 엑셀에 반환
'    - excel_export(파일오픈여부)
'    - 워크북을 만들어서 바탕화면에 저장
'    - 파일오픈이 true면 저장 후 열기
'-----------------------------------------------------------------------------
Sub excel_export(Optional FileOpen As Boolean = False)
    
    Dim tableNM As String, dbNM As String
    Dim strSQL As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    
    '//db명 설정
    tableNM = "accounts.transactions" '//db명.테이블명 - 수정★★
    dbNM = "accounts" '//수정★★
    
    '//DB연결
    connectDB IPAddress, dbNM, "root", DBPassword
    
    '//Select문
    strSQL = "SELECT 열명 FROM " & tableNM & " WHERE 조건식" & ";"
    
    '//SQL문 실행하고 조회된 자료를 레코드셋에 담음
    callDBtoRS strSQL
    If rs.EOF = True Then
        MsgBox "조회 조건에 맞는 자료가 없습니다.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
        
    '//엑셀로 자료 내보내기
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
    
    '//결과보고, 마무리
    fileSNM = Right(fileNM, Len(fileNM) - InStrRev(fileNM, "\"))
    MsgBox "바탕화면에 파일이 생성되었습니다." & vbNewLine & vbNewLine & _
        " - 파일이름: " & fileSNM, vbInformation, banner
    disconnectALL
End Sub

'-----------------------------------------------
'  DB연결
'    - connectDB(서버 IP, 스키마, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'--------------------------
'  DB 및 RS 연결 해제
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
'  레코드셋 설정 및 데이터 반환
'-----------------------------------------------
Sub callDBtoRS(SQLScript As String)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
End Sub

'------------------------------------------------
'  SQL 패턴매칭 검색어 처리('%검색어%')
'------------------------------------------------
Public Function PText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        PText = "'%%'"
    Else
        PText = "'%" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "%'"
    End If
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

'--------------------
'  매크로 최적화
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
'  매크로 최적화 원복
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
'  바탕화면 경로 조회
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

Attribute VB_Name = "A_Common"
Option Explicit

Public Const Banner As String = "프로그램 명칭"
Public Const ODBCDriver As String = "MariaDB ODBC 3.1 Driver"
Public Const programv As String = "Program Version"
Public conn As ADODB.Connection
Public rs As New ADODB.Recordset
Public connIP As String, connDB As String, connUN As String, connPW As String '//Task DB 연결 정보
Public user_id As Integer '사용자코드
Public user_gb As String '사용자구분(SA, AM, MG, WP)
Public user_nm As String '사용자이름
Public checkLogin As Integer '로그인 여부 0: 로그인 안함, 1 = 로그인
Public Const commonPW As String = "Password"
Public cuCode As Integer, pjCode As Integer

'-----------------------
'  Common DB연결
'-----------------------
Sub connectCommonDB()
    connectDB "IP Address", "DB Name", "ID", commonPW
End Sub

'-------------------
'  Task DB연결
'-------------------
Sub connectTaskDB()
    connectDB connIP, connDB, connUN, connPW
End Sub

'-----------------------------------------------
'  DB연결 프로시저
'    - connectDB(서버 IP, 스키마, ID, PW)
'-----------------------------------------------
Sub connectDB(argIP As String, argDB As String, argID As String, argPW As String)
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Driver={" & ODBCDriver & "};Server=" & argIP & ";Port=3306;Database=" & argDB & ";User=" & argID & ";Password=" & argPW & ";Option=2;"
    conn.Open
End Sub

'---------------------------------------------------------------------
'  레코드셋 설정 및 데이터 반환
'    - calDBtoRS(프로시저명, 테이블명, SQL문, 폼이름, 잡이름)
'    - 오류발생 시 에러 핸들링 및 로그 기록
'    - 오류발생 안하면 잡 수행 프로시저에서 로그 기록(필요 시)
'---------------------------------------------------------------------
Sub callDBtoRS(ProcedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional JobNM As String = "데이터 조회")
On Error GoTo ErrHandler

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Source:=SQLScript, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText
    Exit Sub
    
ErrHandler:
    ErrHandle ProcedureNM, tableNM, SQLScript, formNM, JobNM
    writeLog ProcedureNM, tableNM, SQLScript, 1, formNM, JobNM '//오류코드 1
End Sub

'-------------------------------------------------------------------------------------
'  SQL문을 실행하고 실행결과 영향을 받은 레코드 수를 반환
'    - executeSQL(프로시져명, 테이블명, SQL문, 폼이름(옵션), 잡이름(옵션))
'    - SQL문 실행 결과 성공 여부를 알기 위해 영향 받은 레코드 수 검토
'    - 오류발생 시 에러 핸들링 및 로그 기록
'-------------------------------------------------------------------------------------
Public Function executeSQL(ProcedureNM As String, tableNM As String, SQLScript As String, Optional formNM As String = "NULL", Optional JobNM As String = "기타") As Long
On Error GoTo ErrHandler

    Dim affectedCount As Long
    
    conn.Execute CommandText:=SQLScript, recordsaffected:=affectedCount
    executeSQL = affectedCount
    Exit Function
    
ErrHandler:
    ErrHandle ProcedureNM, tableNM, SQLScript, formNM, JobNM
    writeLog ProcedureNM, tableNM, SQLScript, 1, formNM, JobNM '//오류코드 1
End Function

'--------------------------
'  DB 및 RS 연결 해제
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
'  SQL 패턴매칭 검색어 처리('%검색어%')
'------------------------------------------------
Public Function PText(argString As Variant) As String
    If argString = "" Or Len(argString) = 0 Then
        PText = "'%%'"
    Else
        PText = "'%" & Trim(Replace(Replace(argString, "%", "\%"), "'", "''")) & "%'"
    End If
End Function

'---------------------------------------------
'  SQL 스칼라매칭 검색어 처리('검색어')
'---------------------------------------------
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

'----------------
'  전체화면On
'----------------
Sub FullscreenOn()
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
End Sub

'----------------
'  전체화면Off
'----------------
Sub FullscreenOff()
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
End Sub

'---------------------
'  엑셀화면 숨기기
'---------------------
Sub HideExcel()
    Application.Visible = False
End Sub

'---------------------
'  엑셀화면 보이기
'---------------------
Sub ShowExcel()
    Application.Visible = True
End Sub

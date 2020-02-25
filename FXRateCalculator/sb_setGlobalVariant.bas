Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  전역변수 설정
'    - Error 번호 3709, -2147217843
'    - 실행중인 프로시저 재 실행
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional ProcedureNM As String = "NULL")
    Dim strSQL As String
    
    '//전역변수 조회
    connectCommonDB
    strSQL = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "setGlobalVariant", "common.users", strSQL, , "전역변수조회"
    
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저관련 전역변수 설정
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    user_dept = rs("user_dept").Value
    
    disconnectALL
    
    '//오류발생으로 전역변수 재 설정 시 기존 프로시저 실행
    If ProcedureNM <> "NULL" Then Application.Run ProcedureNM
End Sub

'--------------------------------------------------------------------
'  SA가 사용자의 환경 파악을 위해 사용자의 이름으로 로그인
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSQL As String
    
    '//특정 사용자 전역변수 조회
    connectCommonDB
    strSQL = "SELECT * FROM common.users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant2", "common.users", strSQL, "특정사용자전역변수조회"
    
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저코드 설정
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    user_dept = rs("user_dept").Value
    
    disconnectALL
End Sub

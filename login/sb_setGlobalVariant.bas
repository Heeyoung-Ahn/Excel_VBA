Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  Error발생으로 프로젝트 재실행 시 전역변수 재 설정
'    - Error 번호 3709, -2147217843
'    - 실행중인 프로시저 재 실행
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional ProcedureNM As String)
    Dim strSQL As String
    
    connectCommonDB
    
    strSQL = "SELECT * FROM common.v_users WHERE user_id = (SELECT user_id FROM common.v_users WHERE user_nm = " & SText(Application.UserName) & ");"
    
    callDBtoRS "setGlobalVariant", "common.v_users", strSQL
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저코드 설정
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    
    disconnectALL
    If ProcedureNM <> Empty Then Application.Run ProcedureNM
End Sub

'--------------------------------------------------------------------
'  SA가 사용자의 환경 파악을 위해 사용자의 이름으로 로그인
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSQL As String
    
    connectCommonDB
    strSQL = "SELECT * FROM common.v_users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant", "common.v_users", strSQL
    '//작업DB연결을 위한 전역변수 설정
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//유저코드 설정
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    
    disconnectALL
End Sub

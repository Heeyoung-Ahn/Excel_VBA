Attribute VB_Name = "a_WriteLog"
Option Explicit

'----------------------------------------------------------------------------------------------------
'  로그기록
'    - 에러로그: SQL문 실행 시 발생된 로그만 기록(executeSQL, callDBtoRS)
'    - 액션로그: DB에 변경을 프로시저(Insert, Update, Delete) 실행 시 로그 기록
'    - writelog(프로시저명, 테이블명, SQL, 에러코드, 폼이름, 잡이름, 영향받은레코드수)
'-----------------------------------------------------------------------------------------------------
Sub writeLog(ProcedureNM As String, tableNM As String, SQLScript As String, ErrorCD As Integer, Optional formNM As String = "NULL", Optional JobNM As String = "NULL", _
                     Optional affectedCount As Integer = 0)
    Dim strSQL As String
    connectTaskDB
    
    strSQL = "INSERT INTO co_account.logs(procedure_nm, table_nm, sql_script, error_cd, form_nm, job_nm, affectedCount, user_id) " & _
                  "Values(" & SText(ProcedureNM) & ", " & _
                                    SText(tableNM) & ", " & _
                                    SText(SQLScript) & ", " & _
                                    ErrorCD & ", " & _
                                    SText(formNM) & ", " & _
                                    SText(JobNM) & ", " & _
                                    affectedCount & ", " & _
                                    user_id & ");"

    executeSQL "writeLog", "log_table_name", strSQL, formNM, "로그기록"
    disconnectDB
End Sub



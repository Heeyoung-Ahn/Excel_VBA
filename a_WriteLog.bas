Attribute VB_Name = "a_WriteLog"
Option Explicit

'----------------------------------------------------------------------------------------------------
'  로그기록
'    - 에러로그: SQL문 실행 시 발생된 로그만 기록(executeSQL, callDBtoRS)
'    - 액션로그: DB에 변경을 프로시저(Insert, Update, Delete) 실행 시 로그 기록
'    - writelog(프로시저명, 테이블명, SQL, 에러코드, 폼이름, 잡이름, 영향받은레코드수)
'-----------------------------------------------------------------------------------------------------
Sub writeLog(ProcedureNM As String, tableNM As String, SQLScript As String, ErrorCD As Integer, Optional formNM As String = "NULL", Optional JobNM As String = "NULL", _
                     Optional affectedCount As Long = 0)
    Dim strSQL As String
    connectCommonDB
    
    strSQL = "INSERT INTO common.logs(procedure_nm, table_nm, form_nm, job_nm, error_cd, affectedCount, sql_script, user_id) " & _
                  "Values(" & SText(ProcedureNM) & ", " & _
                                    SText(tableNM) & ", " & _
                                    SText(formNM) & ", " & _
                                    SText(JobNM) & ", " & _
                                    ErrorCD & ", " & _
                                    affectedCount & ", " & _
                                    SText(SQLScript) & ", " & _
                                    user_id & ");"

    executeSQL "writeLog", "common.logs", strSQL, , "로그기록"
    disconnectDB
End Sub



Attribute VB_Name = "a_WriteLog"
Option Explicit

'----------------------------------------------------------------------------------------------------
'  �αױ��
'    - �����α�: SQL�� ���� �� �߻��� �α׸� ���(executeSQL, callDBtoRS)
'    - �׼Ƿα�: DB�� ������ ���ν���(Insert, Update, Delete) ���� �� �α� ���
'    - writelog(���ν�����, ���̺���, SQL, �����ڵ�, ���̸�, ���̸�, ����������ڵ��)
'-----------------------------------------------------------------------------------------------------
Sub writeLog(ProcedureNM As String, tableNM As String, SQLScript As String, ErrorCD As Integer, Optional formNM As String = "NULL", Optional JobNM As String = "NULL", _
                     Optional affectedCount As Long = 0)
    Dim strSQL As String
    connectCommonDB
    
    strSQL = "INSERT INTO common.logs(procedure_nm, table_nm, sql_script, error_cd, form_nm, job_nm, affectedCount, user_id) " & _
                  "Values(" & SText(ProcedureNM) & ", " & _
                                    SText(tableNM) & ", " & _
                                    SText(SQLScript) & ", " & _
                                    ErrorCD & ", " & _
                                    SText(formNM) & ", " & _
                                    SText(JobNM) & ", " & _
                                    affectedCount & ", " & _
                                    user_id & ");"

    executeSQL "writeLog", "common.logs", strSQL, , "�αױ��"
    disconnectDB
End Sub


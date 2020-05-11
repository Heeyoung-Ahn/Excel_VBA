Attribute VB_Name = "a_WriteLog"
Option Explicit

'----------------------------------------------------------------------------------------------------
'  �αױ��
'    - �����α�: SQL�� ���� �� �߻��� �α׸� ���(executeSQL, callDBtoRS)
'    - �׼Ƿα�: DB�� ������ ���ν���(Insert, Update, Delete) ���� �� �α� ���
'    - writelog(���ν�����, ���̺��, SQL, �����ڵ�, ���̸�, ���̸�, ����������ڵ��)
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

    executeSQL "writeLog", "common.logs", strSQL, , "�αױ��"
    disconnectDB
End Sub



Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  �������� ����
'    - Error ��ȣ 3709, -2147217843
'    - �������� ���ν��� �� ����
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional ProcedureNM As String = "NULL")
    Dim strSQL As String
    
    '//�������� ��ȸ
    connectCommonDB
    strSQL = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "setGlobalVariant", "common.users", strSQL, , "����������ȸ"
    
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�������� �������� ����
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    user_dept = rs("user_dept").Value
    
    disconnectALL
    
    '//�����߻����� �������� �� ���� �� ���� ���ν��� ����
    If ProcedureNM <> "NULL" Then Application.Run ProcedureNM
End Sub

'--------------------------------------------------------------------
'  SA�� ������� ȯ�� �ľ��� ���� ������� �̸����� �α���
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSQL As String
    
    '//Ư�� ����� �������� ��ȸ
    connectCommonDB
    strSQL = "SELECT * FROM common.users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant2", "common.users", strSQL, "Ư�����������������ȸ"
    
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�����ڵ� ����
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    user_dept = rs("user_dept").Value
    
    disconnectALL
End Sub

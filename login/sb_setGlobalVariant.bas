Attribute VB_Name = "sb_setGlobalVariant"
Option Explicit

'-----------------------------------------------------------------
'  Error�߻����� ������Ʈ ����� �� �������� �� ����
'    - Error ��ȣ 3709, -2147217843
'    - �������� ���ν��� �� ����
'-----------------------------------------------------------------
Sub setGlobalVariant(Optional ProcedureNM As String)
    Dim strSQL As String
    
    connectCommonDB
    
    strSQL = "SELECT * FROM common.v_users WHERE user_id = (SELECT user_id FROM common.v_users WHERE user_nm = " & SText(Application.UserName) & ");"
    
    callDBtoRS "setGlobalVariant", "common.v_users", strSQL
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�����ڵ� ����
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    
    disconnectALL
    If ProcedureNM <> Empty Then Application.Run ProcedureNM
End Sub

'--------------------------------------------------------------------
'  SA�� ������� ȯ�� �ľ��� ���� ������� �̸����� �α���
'--------------------------------------------------------------------
Sub setGlobalVariant2(userNM As String)
    Dim strSQL As String
    
    connectCommonDB
    strSQL = "SELECT * FROM common.v_users WHERE user_nm = " & SText(userNM) & ";"
    callDBtoRS "setGlobalVariant", "common.v_users", strSQL
    '//�۾�DB������ ���� �������� ����
    connIP = rs("argIP").Value
    connDB = rs("argDB").Value
    connUN = rs("argUN").Value
    connPW = rs("argPW").Value
    
    '//�����ڵ� ����
    user_id = rs("user_id").Value
    user_nm = rs("user_nm").Value
    user_gb = rs("user_gb").Value
    
    disconnectALL
End Sub

Attribute VB_Name = "a_ErrHandler"
Option Explicit

'-----------------------------------------------------------------------------------------------------
'  ����ó��: errhandle(���ν�����, ���̺��, SQL��, ���̸�, �۾���)
'    - ���� �߻� ������ ����� �ϱ� ���� �޽��� �ڽ��� ǥ��
'    - ���� �߻��� ���� �α� ����� DB�� ����� ���븸 callDBroRS, executeSQL���� ����
'-----------------------------------------------------------------------------------------------------
Sub ErrHandle(ProcedureNM As String, Optional tableNM As String = "NULL", Optional SQLScript As String = "NULL", Optional formNM As String = "NULL", Optional JobNM As String = "��Ÿ")
    If Err.Number <> 0 Then
        MsgBox "������ �߻��߽��ϴ�." & Space(7) & vbNewLine & _
            " �� ������ �߻��� ������ ĸó�Ͽ� �����ڿ��� �����ּ���." & vbNewLine & vbNewLine & _
            "  �� �۾��� : " & Application.UserName & vbNewLine & _
            "  �� �۾��Ͻ� : " & Now & vbNewLine & _
            "  �� �۾����� : " & JobNM & vbNewLine & vbNewLine & _
            "  �� ���� �߻� vba : " & ProcedureNM & vbNewLine & _
            "  �� ���� �߻� �� : " & formNM & vbNewLine & _
            "  �� ���� �߻� DB : " & tableNM & vbNewLine & _
            "  �� ���� �߻� Script : " & SQLScript & vbNewLine & vbNewLine & vbNewLine & _
            "  �� ���� �ڵ� : " & Err.Number & vbNewLine & _
            "  �� ���� ���� : " & Err.Description & vbNewLine & _
            "  �� ���� �ҽ� : " & Err.Source
    End If
End Sub


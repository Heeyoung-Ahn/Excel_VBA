Attribute VB_Name = "sb_CheckUser"
Option Explicit
Public Const Banner As String = "��ϵ� ����� ����"
Sub OpenEvent()

'#############################################################
'��ũ���� ������ �� ��ϵ� ��������� ���θ� �����ϴ� ���ν���
'#############################################################

Dim aryUser() As Variant
    
    '--//��� ������ ����� ����� �����ϼ���.
    aryUser = Array("�ֿ켮", "������", "������", "�̰���", "����") '--//�ڡ� ������
    
    '--//����� �̸��� �����մϴ�.
    Call sbSetUserName(Application.UserName)
    
    '--//������ ����� �̸��� ��ϵ� ��������� ���θ� �����մϴ�.
    If fnCheckUserName(Application.UserName, aryUser) = False Then
        MsgBox "��ϵ��� ���� ����� �Դϴ�. ���α׷��� �����մϴ�." & vbNewLine & _
                "������ ��û�Ϸ��� ����ڿ��� �����ϼ���.", vbCritical, Banner
        ThisWorkbook.Close False
    Else
        MsgBox "��ϵ� ����� �Դϴ�. �ݰ����ϴ�.", vbInformation, Banner
    End If

    

End Sub

Sub sbSetUserName(UserNM As String)

'###############################
'����� �̸��� �����ϴ� ���ν���
'###############################

    '--//���� ����� �̸��� ����ְ� ��� ������� ���θ� ���´�.
    If MsgBox("���� ����� �̸��� " & UserNM & " �Դϴ�." & vbNewLine & _
                "�ش� �̸��� ��� ����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Banner) = vbNo Then
Back:
        '--//���� ����� �̸��� �ٲٰ� ���� ��� ���ϴ� ���� �Է��Ѵ�.
        UserNM = InputBox("����� ����� �̸��� �Է��ϼ���.", Banner, UserNM)
       
       If UserNM = vbNullString Then GoTo Back
       
       '--//�Է��� ���� ������� �ٽ� �� �� ��� �� ����� �̸��� �ش� ������ ��ȯ�Ѵ�.
        If MsgBox("����� �̸��� " & UserNM & "���� ���� �Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, Banner) = vbYes Then
            Application.UserName = UserNM
        End If
    End If
    
End Sub

'##########################################################################################
'��fnCheckUserName �Լ�
'�����: ���� ����ڰ� ��ϵ� ��������� ���θ� �����Ͽ� �������� ����� ��ȯ�մϴ�.
'���μ�����:
'________UserNM: ������ ����� �̸�
'________aryUser: ��ϵ� ����� �̸����
'##########################################################################################
Function fnCheckUserName(UserNM As String, aryUser As Variant) As Boolean
    Dim var As Variant
    
    '--//�ʱⰪ = False(��ϵ��� ���� �����)
    fnCheckUserName = False
    
    '--//���� ����� �̸��� ��ϵ� ����ڷ� �Ǹ�� ��� ���� True�� ����
    For Each var In aryUser
        If var = UserNM Then
            fnCheckUserName = True
        End If
    Next var

End Function

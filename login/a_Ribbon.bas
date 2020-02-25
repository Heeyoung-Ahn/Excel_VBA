Attribute VB_Name = "a_Ribbon"
Option Explicit

'--------------------
'  �߰���� �޴���
'--------------------
Sub make_menubar()
Call reset_menubar
On Error Resume Next

    With Application.CommandBars("tools").Controls
        With .Add(Type:=msoControlButton)
            .FaceId = 1907
            .Caption = "�α���"
            .OnAction = "LogIn"
        End With
        With .Add(Type:=msoControlButton)
            .FaceId = 5955
            .Caption = "�α׾ƿ�"
            .OnAction = "LogOut"
        End With
        With .Add(Type:=msoControlButton)
            .FaceId = 1088
            .Caption = "���α׷�����"
            .OnAction = "AddinUninstall"
        End With
    End With

On Error GoTo 0
End Sub

'---------------------------------------------------------------------
'  �α���
'    - �߰������ 2�� �̻��� ��� ���ν��� ���� �ٸ��� �ؾ� ��
'---------------------------------------------------------------------
Sub LogIn()
    f_login.Show
End Sub

'------------
'  �α׾ƿ�
'------------
Sub LogOut()
    If checkLogin = 0 Then
        MsgBox Application.UserName & "�� �̹� �α׾ƿ� �Ǿ� �ֽ��ϴ�.", vbInformation, banner
        Exit Sub
    End If
    checkLogin = 0 '�α׾ƿ� ����
    '//�������� �ʱ�ȭ
    connIP = Empty
    connDB = Empty
    connUN = Empty
    connPW = Empty
    user_id = Empty
    user_gb = Empty
    MsgBox "�α׾ƿ� �Ǿ����ϴ�." & Space(7), vbInformation, banner
End Sub

'--------------------------
'  �߰���� �޴��� ����
'--------------------------
Sub reset_menubar()
On Error Resume Next
    Application.CommandBars("WorkSheet Menu Bar").Reset
On Error GoTo 0
End Sub

'------------------
'  �߰���� ����
'------------------
Sub AddinUninstall()
    reset_menubar
    ThisWorkbook.Close False
End Sub


Attribute VB_Name = "a_Ribbon"
Option Explicit

'-----------------------------------------------------
'  ���� �޴��� Button ID�� ���� ó�� ���ν���
'-----------------------------------------------------
Sub run_RibbonControl(Button As Office.IRibbonControl)
    Select Case Button.ID
        '//���α׷�
        Case "FX_Calculator":     Call FX_Calculator
        Case "InsertPicture":    Call InsertPicture
        Case "InsertPicture2":    Call InsertPicture2
                      
        '//����
        Case "LogIn":     Call LogIn
        Case "LogOut":     Call LogOut
        Case "AddinUninstall":     Call AddinUninstall
        
        Case Else:     Call RibbonButton_Error(Button.ID)
    End Select
End Sub

'-------------------------------------------------------------------
'  Button ID�� ���� ó�� ���ν����� ���� ��� ���� �޽���
'-------------------------------------------------------------------
Sub RibbonButton_Error(sbID As String)
   MsgBox "�����Ͻ� �޴�(" & sbID & ")�� ���� �غ� �Ǿ� ���� �ʽ��ϴ�.", vbCritical, banner
End Sub

'-----------
'  �α���
'-----------
Sub LogIn()
    If checkLogin = 1 Then
        MsgBox Application.UserName & "�� �̹� �α��� �Ǿ� �ֽ��ϴ�.", vbInformation, banner
        Exit Sub
    End If
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
    user_dept = Empty
    MsgBox "�α׾ƿ� �Ǿ����ϴ�." & Space(7), vbInformation, banner
End Sub

'------------------------------------
'  ���� ������ �ݾ� ���� �� ����
'------------------------------------
Sub AddinUninstall()
   ThisWorkbook.Close False
End Sub

'---------------
'  ȯ����ȸ��
'---------------
Sub FX_Calculator()
    f_currency_cal.Show vbModeless
End Sub


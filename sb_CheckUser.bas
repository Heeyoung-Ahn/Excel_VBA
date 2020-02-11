Attribute VB_Name = "sb_CheckUser"
Option Explicit
Public Const banner As String = "����VBA����(�����üũ)"
Dim registeredUser As Variant

'------------------------------
'  ����� �̸� ����
'  ����� �̸� ��� Ȯ��
'------------------------------
Sub Workbook_Open()
    
    '//���α׷� ����� ����
    registeredUser = Array("����", "�����2", "�����3")
    
    '//���α׷� ����� �̸� ����
    Call setUserName(Application.UserName)
    
    '//����� �̸� ��� ���� Ȯ��: ��ϵ��� ���� ����ڴ� ��ũ�� ����
    If checkUserName(Application.UserName, registeredUser) = False Then
        MsgBox "'" & Application.UserName & "'���� ���α׷� ����ڷ� ��ϵǾ� ���� �ʽ��ϴ�." & vbNewLine & _
            "���� ������ �����մϴ�.", vbCritical, banner
        ThisWorkbook.Close savechanges:=False
    Else
        MsgBox "'" & Application.UserName & "'���� ���α׷� ����ڷ� Ȯ�εǾ����ϴ�." & vbNewLine & _
            "�� ������ ���������� ����� �� �ֽ��ϴ�.", vbInformation, banner
    End If
End Sub

'------------------------
'  ����� �̸� ����
'------------------------
Sub setUserName(userNM As String)

    If MsgBox("�� ���α׷����� ����� ������� �̸��� " & userNM & "���� �ϰڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Do
            Application.UserName = InputBox("������� �̸��� �Է��� �ּ���.", banner, Application.UserName)
            userNM = Application.UserName
        Loop Until userNM <> vbNullString
    End If
    MsgBox "�� ���α׷� ������� �̸��� '" & userNM & "'���� �����Ͽ����ϴ�.", vbInformation, banner

End Sub

'------------------------
'  ��ϵ� ����� üũ
'------------------------
Function checkUserName(argUserNM As String, argRegisteredUser As Variant) As Boolean
    Dim userNM As Variant
       
    checkUserName = False
    For Each userNM In argRegisteredUser
        If userNM = argUserNM Then
            checkUserName = True
            Exit For
        End If
    Next userNM
    
End Function


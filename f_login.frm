VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_login 
   Caption         =   "�α���"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   OleObjectBlob   =   "f_login.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "f_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------
'  �α���â ���� �� �α��ΰ���
'-----------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    If checkLogin = 0 Then
        MsgBox "�α��� ������ Ȯ�ε��� �ʾҽ��ϴ�." & Space(7) & vbNewLine & _
            "���α׷��� �����մϴ�.", vbInformation, Banner
        ThisWorkbook.Close savechanges:=False
    End If
End Sub

'------------------------------------------------------
'  �α��� ��(common)
'  - ���α׷� ���� üũ�Ͽ� ���� ���� ��� ����
'  - IP���
'------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim strSQL As String
    
    '//���ʼ���
    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    Me.Caption = Banner
    txt1.Value = Application.UserName
    Me.lbl_pv = programv
    Me.lbl_report = reportfile_nm
        
    '//��ϵ� ����� üũ
    If checkUserNm(txt1.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� ������ �ּ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�� �ּ���.", vbInformation, Banner
        GoTo n
    End If
    
    '//��й�ȣ ���� ���� üũ
    Call checkInitialPW
n:
    txt2.SetFocus
    Exit Sub
ErrHandler:
    End
End Sub

'-------------------------------------------------------------------------------------
'  ��ϵ� ����� üũ
'    - txt1�� �Էµ� ����ڰ� ��ϵ� ��������� �����Ͽ� true / false �� ��ȯ
'-------------------------------------------------------------------------------------
Private Function checkUserNm(ByVal argUserNM As String) As Boolean
    Dim strSQL As String
    
    connectCommonDB
    strSQL = "SELECT * FROM common.v_users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    callDBtoRS "checkUserNm", "common.v_users", strSQL, "f_login"
    
    If rs.RecordCount = 0 Then
        checkUserNm = False
    Else
        checkUserNm = True
    End If
    
    disconnectALL
End Function

'---------------------------------------
'  txt1���� exit �� ���
'    - ����� �̸� ��Ͽ��� üũ
'    - ��й�ȣ �ʱ� ���� ���� üũ
'---------------------------------------
Private Sub txt1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt1 = Empty Then
        Exit Sub
    End If
    
    '//����� �̸� ��� ���� üũ
    If checkUserNm(txt1.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� ������ �ּ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�� �ּ���.", vbInformation, Banner
        txt1.SetFocus
        Exit Sub
    End If
    If txt1.Value <> Application.UserName Then
        Application.UserName = txt1.Value
    End If
    
    '//��й�ȣ �ʱ� ���� ���� üũ
    Call checkInitialPW
    
End Sub

'----------------------------------------------------------------------------------------
'  ��ϵ� ������� ��� ��й�ȣ�� �����Ǿ� �־����� üũ�ϰ� �����ϵ��� ����
'----------------------------------------------------------------------------------------
Private Sub checkInitialPW()
    Dim strSQL As String
    Dim strPW As Integer
    Dim user_pw As Variant
    Dim affectedCount As Long
    
    connectCommonDB
    strSQL = "SELECT pw_initialize FROM common.v_users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    callDBtoRS "checkInitialPW", "common.v_users", strSQL, "f_login"
    
    strPW = rs("pw_initialize").Value
    If strPW = 1 Then '//PW �Է� �̷��� ������ PW ����
        MsgBox "��й�ȣ�� �����Ǿ� ���� �ʽ��ϴ�.", vbInformation, Banner
        registerNewPW
    End If
    disconnectALL
End Sub

'-----------------------
'  �űԺ�й�ȣ ���
'-----------------------
Private Sub registerNewPW()
    Dim strSQL As String
    Dim strPW As Integer
    Dim user_pw As Variant
    Dim affectedCount As Long
    '��й�ȣ �Է� �ޱ�
    Do
        user_pw = InputBoxPW("�ű� ��й�ȣ�� ��ҹ��ڸ� �����Ͽ� 4�ڸ� �̻����� ������ �ּ���.", Banner)
    Loop Until user_pw <> Empty And Len(user_pw) > 3
    '��й�ȣ ���
    strSQL = "UPDATE common.users SET user_pw = SHA2(" & SText(user_pw) & ", 512) WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    affectedCount = executeSQL("checkInitialPW", "common.users", strSQL, "f_login", "�ʱ��й�ȣ����")
    '��й�ȣ �ʱ�ȭ ��Ȱ��ȭ
    strSQL = "UPDATE common.users SET pw_initialize = 0 WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    affectedCount = executeSQL("checkInitialPW", "common.users", strSQL, "f_login", "�ʱ��й�ȣ����")
    If affectedCount = 0 Then
        MsgBox "��й�ȣ�� �������� �ʾҽ��ϴ�." & Space(7) & vbNewLine & _
            "�����ڿ��� �����Ͽ� �ֽñ� �ٶ��ϴ�.", vbInformation, Banner
        'ThisWorkbook.Close savechanges:=False
    Else
        MsgBox "��й�ȣ ������ �Ϸ�Ǿ����ϴ�." & Space(7), vbInformation, Banner
    End If
    disconnectALL
End Sub

'---------------------------------------
'  Ȯ�ι�ư ��
'    - ����� �̸� ��� ���� üũ
'    - ���α׷� �ֽŹ��� Ȯ��
'    - IPüũ
'    - ��й�ȣ �´� �� ����
'    - ����ȸ�跱ó ����Ű ����: ALT + ���ʹ���Ű
'    - ȯ���λ�
'---------------------------------------
Private Sub cmd_query_Click()
    Dim strSQL As String
    Dim affectedCount As Long
    Dim ipRng As Integer
    
    '//����� �̸� ��� ���� üũ
    If txt1 = Empty Then
        MsgBox "������� �̸��� �Է��ϼ���.", vbInformation, Banner
        Exit Sub
    End If
    If checkUserNm(txt1.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� ������ �ּ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�� �ּ���.", vbInformation, Banner
        txt1.SetFocus
        Exit Sub
    End If
    If txt1.Value <> Application.UserName Then
        Application.UserName = txt1.Value
    End If
    
    '//��й�ȣ �Է� ���� üũ
    If txt2 = Empty Then
        MsgBox "��й�ȣ�� �Է��ϼ���.", vbInformation, Banner
        txt2.SetFocus
        Exit Sub
    End If
    
    '//���α׷� ���� Ȯ��
    strSQL = "SELECT programv FROM common.users WHERE user_id = 3"
    connectCommonDB
    callDBtoRS "txt1_Exit", "common.users", strSQL, Me.Name, "���α׷����� Ȯ��"
    If rs("programv").Value <> programv Then
        MsgBox "����Ϸ��� ����ȸ�����α׷��� �ֽŹ����� �ƴմϴ�." & vbNewLine & _
                     "���α׷� ���� ������ ���� �ֽŹ������� ����� �ּ���.", vbInformation, Banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//IPȮ��
    ipRng = Mid(GetLocalIPaddress, InStr(5, GetLocalIPaddress, ".") + 2, 2)
    If ipRng <> 10 And ipRng <> 11 Then
        MsgBox "����ȸ�����α׷��� ���� PC������ ��� �����մϴ�." & vbNewLine & _
                     "���α׷��� �����մϴ�.", vbInformation, Banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//��й�ȣ �´� �� ����
    If checkPW(txt2.Value) = True Then
        '��й�ȣ�� ������ Welcome
        checkLogin = 1
        setGlobalVariant
        '//����ȸ�� ��ó ����Ű ����
        Application.OnKey "%{LEFT}", "start_co_account"
        '//DB�� IP�Է�
        '[����IP�����]
        strSQL = "UPDATE common.users SET user_ip = NULL WHERE user_id = " & user_id & ";"
        connectCommonDB
        affectedCount = executeSQL("cmd_query_Click", "common.users", strSQL, Me.Name, "�����IP���")
        '[�ű�IP�ֱ�]
        strSQL = "UPDATE common.users SET user_ip = " & SText(GetLocalIPaddress) & " WHERE user_id = " & user_id & ";"
        affectedCount = executeSQL("cmd_query_Click", "common.users", strSQL, Me.Name, "�����IP���")
        If affectedCount > 0 Then
            writeLog "cmd_query_Click", "common.users", strSQL, 0, Me.Name, "�����IP���", affectedCount
        End If
        '//ȯ���λ�
        MsgBox Application.UserName & "�� ������ ��������." & Space(7) & vbNewLine & vbNewLine & _
                 "������ " & Format(Date, "YYYY-MM-DD") & "�� �Դϴ�." & vbNewLine & _
                 "���õ� ANIMO!", vbInformation, Banner
        today = Date
        Unload Me
    Else
        '��й�ȣ�� �ٸ��� �ٽ� �Է�
        MsgBox "��й�ȣ�� Ʋ�Ƚ��ϴ�." & Space(7) & vbNewLine & _
            "��й�ȣ�� �ٽ� �Է��Ͽ� �ּ���.", vbInformation, Banner
        txt2.Value = Empty
        txt2.SetFocus
        Exit Sub
    End If
        
End Sub

'------------------------------------------------------------------------
'  �Էµ� ��й�ȣ�� �´��� Ʋ���� �����Ͽ� true / false �� ��ȯ
'------------------------------------------------------------------------
Private Function checkPW(ByVal argPW As String) As Boolean
    Dim strSQL As String
    Dim strPW As Variant
    
    connectCommonDB
    strSQL = "SELECT user_pw FROM common.v_users WHERE user_id = (SELECT user_id FROM common.v_users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "checkPW", "common.v_users", strSQL, "f_login"
    
    strPW = rs("user_pw").Value
    If strPW <> to_SHA512(argPW) Then
        checkPW = False
    Else
        checkPW = True
    End If
End Function

Private Sub cmd_close_Click()
    Unload Me
End Sub

'---------------------------------------
'  ��й�ȣ ����
'    - ���� ��й�ȣ Ȯ��
'    - �ű� ��й�ȣ �Է�
'---------------------------------------
Private Sub cmd_chgPW_Click()
    Dim oldPW As String
    Dim newPW As String
    If MsgBox("��й�ȣ�� �����ϰڽ��ϱ�?", vbQuestion + vbYesNo, Banner) = vbNo Then
        Exit Sub
    Else
        oldPW = InputBoxPW("���� ��й�ȣ�� �Է��ϼ���.", Banner)
        If StrPtr(oldPW) = 0 Then
            MsgBox "���� ��й�ȣ �Է��� ��ҵǾ����ϴ�.", vbInformation, Banner
            Exit Sub
        Else
            If checkPW(oldPW) = True Then
                registerNewPW
            Else
                MsgBox "���� ��й�ȣ�� ��ġ���� �ʽ��ϴ�." & vbNewLine & _
                             "�����ڿ��� �����Ͽ� �ּ���.", vbInformation, Banner
            End If
        End If
    End If
End Sub


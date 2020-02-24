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

'------------------------------------
'  �α���â ���� �� �α��ΰ���
'------------------------------------
Private Sub UserForm_Terminate()
    If checkLogin = 0 Then
        MsgBox "�α��� ������ Ȯ�ε��� �ʾҽ��ϴ�." & Space(7) & vbNewLine & _
            "���α׷��� �����մϴ�.", vbInformation, banner
        reset_menubar
        ThisWorkbook.Close savechanges:=False
    End If
    disconnectALL
End Sub

'------------------------------------------------------
'  �α��� ��(common)
'  - ID, PWüũ
'  - ���α׷� ���� üũ
'  - IPüũ
'------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim strSQL As String
    
    '//���ʼ���
    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    Me.Caption = banner
    txt_ID.Value = Application.UserName
    Me.lbl_pv = programv
        
    '//��ϵ� ����� üũ
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� ������ �ּ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�� �ּ���.", vbInformation, banner
        GoTo n
    End If
    
    '//��й�ȣ ���� ���� üũ
    Call checkInitialPW
n:
    txt_PW.SetFocus
    Exit Sub
ErrHandler:
    End
End Sub

'---------------------------------------------------------------------------------------
'  ��ϵ� ����� üũ
'    - txt_ID�� �Էµ� ����ڰ� ��ϵ� ��������� �����Ͽ� true / false �� ��ȯ
'---------------------------------------------------------------------------------------
Private Function checkUserNm(ByVal argUserNM As String) As Boolean
    Dim strSQL As String
    
    connectCommonDB
    strSQL = "SELECT * FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "checkUserNm", "common.users", strSQL, "f_login", "�����Ȯ��"
    
    If rs.RecordCount = 0 Then
        checkUserNm = False
    Else
        checkUserNm = True
    End If
    
    disconnectALL
End Function

'---------------------------------------
'  txt_ID���� exit �� ���
'    - ����� �̸� ��Ͽ��� üũ
'    - ��й�ȣ �ʱ� ���� ���� üũ
'---------------------------------------
Private Sub txt_ID_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt_ID = Empty Then
        Exit Sub
    End If
    
    '//����� �̸� ��� ���� üũ
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� ������ �ּ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�� �ּ���.", vbInformation, banner
        txt_ID.SetFocus
        Exit Sub
    End If
    If txt_ID.Value <> Application.UserName Then
        Application.UserName = txt_ID.Value
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
    
    connectCommonDB
    strSQL = "SELECT pw_initialize FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "checkInitialPW", "common.users", strSQL, "f_login", "����ں�й�ȣ��Ͽ�����ȸ"
    
    strPW = rs("pw_initialize").Value
    If strPW = 1 Then '//PW �Է� �̷��� ������ PW ����
        MsgBox "��й�ȣ�� �����Ǿ� ���� �ʽ��ϴ�." & vbNewLine & _
                     "�����ȣ ����ȭ������ �̵��մϴ�.", vbInformation, banner
        Call registerNewPW
    End If
    disconnectALL
End Sub

'--------------------------------------------------------------------------------------------------------------------
'  ����ڴ� ��ϵ� ����������� ��й�ȣ ������ �ȵǾ� �ִ� ���(pw_initialize = 1) �űԺ�й�ȣ ���
'--------------------------------------------------------------------------------------------------------------------
Private Sub registerNewPW()
    Dim strSQL As String
    Dim strPW As Integer
    Dim user_pw As Variant
    Dim affectedCount As Long
    
    '//��й�ȣ �Է� �ޱ�
    Do
        user_pw = InputBoxPW("�ű� ��й�ȣ�� ��ҹ��ڸ� �����Ͽ� 4�ڸ� �̻����� ������ �ּ���.", banner)
    Loop Until user_pw <> Empty And Len(user_pw) > 3
    
    '//��й�ȣ ���(��ȣȭ)
    connectCommonDB
    strSQL = "UPDATE common.users SET user_pw = SHA2(" & SText(user_pw) & ", 512) WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    affectedCount = executeSQL("registerNewPW", "common.users", strSQL, "f_login", "�ʱ��й�ȣ����")
    If affectedCount > 0 Then
         writeLog "registerNewPW", "common.users", strSQL, 0, Me.Name, "�����PW���", affectedCount
    End If
    disconnectALL
    
    '//��й�ȣ ��� Ȯ��
    If affectedCount = 0 Then
        MsgBox "��й�ȣ�� �������� �ʾҽ��ϴ�." & Space(7) & vbNewLine & _
            "�����ڿ��� �����Ͽ� �ֽñ� �ٶ��ϴ�.", vbInformation, banner
    Else
         '//��й�ȣ �ʱ�ȭ ��Ȱ��ȭ
         connectCommonDB
        strSQL = "UPDATE common.users SET pw_initialize = 0 WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
        executeSQL "registerNewPW", "common.users", strSQL, "f_login", "��й�ȣ�ʱ�ȭ��Ȱ��ȭ"
        writeLog "registerNewPW", "common.users", strSQL, 0, Me.Name, "��й�ȣ�ʱ�ȭ��Ȱ��ȭ", 1
        MsgBox "��й�ȣ ������ �Ϸ�Ǿ����ϴ�." & Space(7), vbInformation, banner
    End If
    disconnectALL
End Sub

'--------------------------------------
'  Ȯ�ι�ư ��
'    - ����� �̸� ��� ���� üũ
'    - ���α׷� �ֽŹ��� Ȯ��
'    - IPüũ
'    - ��й�ȣ üũ
'    - ȯ���λ�
'---------------------------------------
Private Sub cmd_query_Click()
    Dim strSQL As String
    Dim affectedCount As Long
    Dim ipRng As Integer
    
    '//����� �̸� ��� ���� üũ
    If txt_ID = Empty Then
        MsgBox "����� �̸��� �Է��ϼ���.", vbInformation, banner
        Exit Sub
    End If
    If checkUserNm(txt_ID.Value) = False Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�." & Space(7) & vbNewLine & _
                "�α��� â���� �̸��� �����ϼ���." & Space(7) & vbNewLine & _
                "����� ����� �ʿ��ϸ� �����ڿ��� ��û�ϼ���.", vbInformation, banner
        txt_ID.SetFocus
        Exit Sub
    End If
    If txt_ID.Value <> Application.UserName Then
        Application.UserName = txt_ID.Value
    End If
    
    '//��й�ȣ �Է� ���� üũ
    If txt_PW = Empty Then
        MsgBox "��й�ȣ�� �Է��ϼ���.", vbInformation, banner
        txt_PW.SetFocus
        Exit Sub
    End If
    
    '//���α׷� ���� Ȯ��
    connectCommonDB
    strSQL = "SELECT programv FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "cmd_query_Click", "common.users", strSQL, Me.Name, "���α׷����� Ȯ��"
    If UCase(rs("programv").Value) <> UCase(programv) Then
        MsgBox "����Ϸ��� ���α׷��� �ֽŹ����� �ƴմϴ�." & vbNewLine & _
                     "���α׷� ���� ������ ���� �ֽŹ������� ����� �ּ���.", vbInformation, banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//IPȮ��
    'IP�Է� ���� Ȯ��
    strSQL = "SELECT user_ip FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
    callDBtoRS "cmd_query_Click", "common.users", strSQL, "f_login", "�����IPȮ��"
    
    If IsNull(rs("user_ip").Value) Then '���� �����̸� IP ���
        If MsgBox("������ PC�� ������� PC�� ����մϴ�." & vbNewLine & _
                         "��ϵ� PC�� �ٸ� PC������ ���α׷� ����� ���ѵ˴ϴ�." & vbNewLine & _
                         "�����ϰڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
            disconnectALL
            Exit Sub
        Else
            '[�ű�IP�ֱ�]
            strSQL = "UPDATE common.users SET user_ip = " & SText(GetLocalIPaddress) & " WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
            executeSQL "cmd_query_Click", "common.users", strSQL, Me.Name, "�����IP���"
            writeLog "cmd_query_Click", "common.users", strSQL, 0, Me.Name, "�����IP���", 1
        End If
    Else '���� ���� �ƴ� ��� IP üũ
        strSQL = "SELECT user_ip FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt_ID.Value) & ");"
        callDBtoRS "cmd_query_Click", "common.users", strSQL, Me.Name, "�����IPȮ��"
        If rs("user_ip").Value <> GetLocalIPaddress Then
            MsgBox "�� ���α׷��� ���� PC������ ��� �����մϴ�." & vbNewLine & _
                         "������� PC ��� ������ �ʿ��ϸ� �����ڿ��� ��û�ϼ���." & vbNewLine & _
                         "���α׷��� �����մϴ�.", vbInformation, banner
            disconnectALL
            cmd_close_Click
        End If
    End If
    
    '//��й�ȣ �´� �� ����
    If checkPW(txt_PW.Value) = True Then
        '�α��� �� 1, �۷ι� ���� ����
        checkLogin = 1
        setGlobalVariant
        'ȯ���λ�
        MsgBox Application.UserName & "�� ������ ��������." & Space(7) & vbNewLine & vbNewLine & _
                 "������ " & Format(Date, "YYYY-MM-DD") & "�� �Դϴ�." & vbNewLine & _
                 "���õ� ANIMO!", vbInformation, banner
        'today�� ���� ��¥ ����
        today = Date
        Unload Me
    Else
        '��й�ȣ�� �ٸ��� �ٽ� �Է�
        MsgBox "��й�ȣ�� Ʋ�Ƚ��ϴ�." & Space(7) & vbNewLine & _
            "��й�ȣ�� �ٽ� �Է��Ͽ� �ּ���.", vbInformation, banner
        txt_PW.Value = Empty
        txt_PW.SetFocus
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
    strSQL = "SELECT user_pw FROM common.users WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Application.UserName) & ");"
    callDBtoRS "checkPW", "common.users", strSQL, "f_login"
    
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
    
    If MsgBox("��й�ȣ�� �����ϰڽ��ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    Else
        oldPW = InputBoxPW("���� ��й�ȣ�� �Է��ϼ���.", banner)
        If StrPtr(oldPW) = 0 Then
            MsgBox "���� ��й�ȣ �Է��� ��ҵǾ����ϴ�.", vbInformation, banner
            Exit Sub
        Else
            If checkPW(oldPW) = True Then
                registerNewPW
            Else
                MsgBox "���� ��й�ȣ�� ��ġ���� �ʽ��ϴ�." & vbNewLine & _
                             "�����ڿ��� �����Ͽ� �ּ���.", vbInformation, banner
            End If
        End If
    End If
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_unit_converter 
   Caption         =   "������ȯ��"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6120
   OleObjectBlob   =   "f_unit_converter.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "f_unit_converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------
'  �� ���� ��
'--------------
Private Sub UserForm_Terminate()

End Sub

'-------------------------------
'  ������ȯ ��(common)
'-------------------------------
Private Sub UserForm_Initialize()
    Dim strSQL As String

    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    
    '//���ʼ���
'    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    
    '//�ؽ�Ʈ�ڽ�
    txt1.Value = 1
    
    '//���� ���� �޺��ڽ�
    With Me.cbo_unit
        .ColumnCount = 2
        .ColumnHeads = False
        .ColumnWidths = "0,108"
        .TextColumn = 2
        .ListWidth = "108"
        .TextAlign = fmTextAlignCenter
        .IMEMode = fmIMEModeHangul
        .Style = fmStyleDropDownCombo
    End With
    strSQL = "SELECT DISTINCT a.unit_gb, a.unit_gb_ko FROM co_account.unit_type a WHERE a.suspended = 0 ORDER BY a.sort_order;"
    loadDataToCBox Me.cbo_unit, strSQL, "co_account.unit_type", Me.Name
    cbo_unit.ListIndex = 0
    
    With txt2
        .Locked = True
        .BackColor = &H80000018
    End With
        
    txt1.SetFocus
End Sub

'-----------------------
'  ���� �޺��ڽ� ����
'-----------------------
Private Sub make_cbo_unit1()
    Dim strSQL As String
    '//����1 �޺��ڽ�
    With Me.cbo_unit1
        .ColumnCount = 3
        .ColumnHeads = False
        .ColumnWidths = "0,108,0"
        .TextColumn = 2
        .ListWidth = "108"
        .TextAlign = fmTextAlignCenter
        .IMEMode = fmIMEModeHangul
        .Style = fmStyleDropDownCombo
    End With
    strSQL = "SELECT a.unit_id, a.unit, a.unit_standard FROM co_account.unit_type a WHERE a.suspended = 0 AND a.unit_gb = " & SText(cbo_unit.Column(0, cbo_unit.ListIndex)) & " ORDER BY a.sort_order;"
    loadDataToCBox Me.cbo_unit1, strSQL, "co_account.unit_type", Me.Name
    cbo_unit1.ListIndex = 0
End Sub
Private Sub make_cbo_unit2()
    Dim strSQL As String
    '//����2 �޺��ڽ�
    With Me.cbo_unit2
        .ColumnCount = 3
        .ColumnHeads = False
        .ColumnWidths = "0,108,0"
        .TextColumn = 2
        .ListWidth = "108"
        .TextAlign = fmTextAlignCenter
        .IMEMode = fmIMEModeHangul
        .Style = fmStyleDropDownCombo
    End With
    strSQL = "SELECT a.unit_id, a.unit, a.unit_standard FROM co_account.unit_type a WHERE a.suspended = 0 AND a.unit_gb = " & SText(cbo_unit.Column(0, cbo_unit.ListIndex)) & " ORDER BY a.sort_order;"
    loadDataToCBox Me.cbo_unit2, strSQL, "co_account.unit_type", Me.Name
    cbo_unit2.ListIndex = 1
End Sub

'----------------------
'  ī�װ� �ٲ��
'    - ����1 ����
'    - ����2 ����
'----------------------
Private Sub cbo_unit_Change()
    Me.txt1.Value = 1
    Call make_cbo_unit1
    Call make_cbo_unit2
    txt1.SetFocus
End Sub

'----------------------
'  ������ȯ
'----------------------
Private Sub txt1_Change()
    If cbo_unit1.ListIndex <> -1 And cbo_unit2.ListIndex <> -1 Then
        Call cmd_refer_Click
    End If
End Sub
Private Sub txt1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Call adjust_decimal(Me.txt1, txt1.Value)
End Sub
Private Sub cbo_unit1_Change()
    If cbo_unit2.ListIndex <> -1 Then Call cmd_refer_Click
End Sub
Private Sub cbo_unit2_Change()
    If cbo_unit1.ListIndex <> -1 Then Call cmd_refer_Click
End Sub
Private Sub cmd_refer_Click()
    Dim unit1S As Boolean, unit2S As Boolean '��ȯ�Ϸ��� ������ �⺻�������� Ȯ��
    Dim amt1 As Double, amt2 As Double '��ȭ�Ϸ��� ���� 2�� ��� �⺻������ �ƴ� ��� ������ ���� / currency�� ������������ ������ ��� �Ҽ��� 4�ڸ��ۿ� �ȵǼ� double�� ����
    Dim strSQL As String
    
    '//�Է°���: checkTextBox(�ؽ�Ʈ�ڽ� �̸�, �ؽ�Ʈ�ڽ� Ÿ��Ʋ, �ʼ�����, ��������, ���� ����, �Է� �� ��Ŀ��)
    If Me.txt1 = Empty Then
        MsgBox "������ ��ȯ�� ��ġ�� �Է��ϼ���.", vbInformation, Banner
        Exit Sub
    Else
        If checkTextBox(Me.txt1, "������ȯ�� ��ġ", True, "NUMERIC", , True) = False Then Exit Sub
    End If
    If cbo_unit1.ListIndex = -1 Then
        MsgBox "��ȯ�� ������ �Է��ϼ���.", vbInformation, Banner
        Exit Sub
    End If
    If cbo_unit2.ListIndex = -1 Then
        MsgBox "��ȯ�� ������ �Է��ϼ���.", vbInformation, Banner
        Exit Sub
    End If
    
    '//������ȯ
        '�⺻���� ���� Ȯ��
        If cbo_unit1.Column(0, cbo_unit1.ListIndex) = cbo_unit1.Column(2, cbo_unit1.ListIndex) Then
            unit1S = True
        Else
            unit1S = False
        End If
        If cbo_unit2.Column(0, cbo_unit2.ListIndex) = cbo_unit2.Column(2, cbo_unit2.ListIndex) Then
            unit2S = True
        Else
            unit2S = False
        End If
        '���̽��� ������ȯ
        If (cbo_unit1.Column(0, cbo_unit1.ListIndex) = cbo_unit2.Column(0, cbo_unit2.ListIndex)) Then '�� ������ ������ ���
            amt1 = 1
        ElseIf unit1S = True And unit2S = False Then '����1�� �⺻����, ����2�� �ƴ� ���
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & " AND a.unit_id2 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "������ȯ����ȸ"
            If rs.EOF <> True Then
                amt1 = rs("value").Value
            End If
            disconnectALL
        ElseIf unit1S = False And unit2S = True Then '����2�� �⺻������ ���
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id2 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & " AND a.unit_id1 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "������ȯ����ȸ"
            If rs.EOF <> True Then
                amt1 = 1 / rs("value").Value
            End If
            disconnectALL
        ElseIf unit1S = False And unit2S = False Then '�� �� �⺻������ �ƴ� ���
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit1.Column(2, cbo_unit1.ListIndex) & " AND a.unit_id2 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "������ȯ����ȸ"
            If rs.EOF <> True Then
                amt1 = 1 / rs("value").Value
            End If
            disconnectALL
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit2.Column(2, cbo_unit2.ListIndex) & " AND a.unit_id2 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "������ȯ����ȸ"
            If rs.EOF <> True Then
                amt2 = rs("value").Value
            End If
            disconnectALL
            amt1 = amt1 * amt2
        End If
    Call adjust_decimal(Me.txt2, Round(txt1.Value * amt1, 10)) 'decimal�� 10�ڸ������� ©�� ����
End Sub

'----------------------------
'  �Ҽ��� �ڸ��� ���� ��ȯ
'----------------------------
Private Sub adjust_decimal(argTB As MSForms.TextBox, argValue As Double)
    Dim noA As Integer '��ü ���ڼ�
    Dim noB As Integer '������ '0'���� ��ġ
    Dim noC As Integer '�Ҽ��� �ڸ���
    
    noA = Len(Format(argValue, "@")) '������ ��� len�Լ��� �ȸԾ ���������� ��ȯ�Ͽ� ����
    noB = InStrRev(argValue, "0")
    If InStr(argValue, ".") = 0 Then
        noC = 0
        argTB.Value = Format(argValue, "#,##0")
    Else
        noC = noA - InStr(argValue, ".")
        Do While noA = noB
            argValue = Left(argValue, noB - 1)
            
            noA = Len(Format(argValue, "@"))
            noB = InStrRev(argValue, "0")
        Loop
        noC = noA - InStr(argValue, ".")
        Select Case noC
            Case 10
                argTB.Value = Format(argValue, "#,##0.0000000000")
            Case 9
                argTB.Value = Format(argValue, "#,##0.000000000")
            Case 8
                argTB.Value = Format(argValue, "#,##0.00000000")
            Case 7
                argTB.Value = Format(argValue, "#,##0.0000000")
            Case 6
                argTB.Value = Format(argValue, "#,##0.000000")
            Case 5
                argTB.Value = Format(argValue, "#,##0.00000")
            Case 4
                argTB.Value = Format(argValue, "#,##0.0000")
            Case 3
                argTB.Value = Format(argValue, "#,##0.000")
            Case 2
                argTB.Value = Format(argValue, "#,##0.00")
            Case 1
                argTB.Value = Format(argValue, "#,##0.0")
            Case 0
                argTB.Value = Format(argValue, "#,##0")
        End Select
    End If
End Sub

'-----------------------------------------
'  �� �ݱ�
'-----------------------------------------
Private Sub cmd_close_Click()
    Unload Me
End Sub

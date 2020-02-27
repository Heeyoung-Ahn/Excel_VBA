VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_currency_cal 
   Caption         =   "ȯ����ȸ��"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8640
   OleObjectBlob   =   "f_currency_cal.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "f_currency_cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//����Ʈ �÷� ��
Dim TB1 As String '//���� ����� DB ���̺�
Const strOrderBy As String = "currency_un ASC" '//DB���� sort_order �ʵ�
Dim caseSave As Integer '//1: �߰�, 2: ����, 3: ����(SUSPEND), 4: ��������
Dim queryKey As Integer '//����Ʈ ��ġ ��ȯ�� ���� id

'--------------
'  �� ���� ��
'--------------
Private Sub UserForm_Terminate()

End Sub

'-------------------------------
'  ȯ����� ��(common)
'-------------------------------
Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//�������� �缳��
    If checkLogin = 0 Then f_login.Show '//�α���üũ
    
    '//���ʼ���
    cntlst1Col = 5 '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��, ���ļ���
'    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    
    '//���� ����� object ����
    TB1 = "fx_calculator.currency_cal"
    
    With txtC3
        .Locked = True
        .BackColor = &H80000018
    End With
        
    '//����Ʈ���� ����
    With lst1
        .ColumnCount = cntlst1Col
        .ColumnHeads = False
        .ColumnWidths = "0,48,70,70,70" 'ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        .Width = 260
        .TextAlign = fmTextAlignLeft
        .Font = "���� ���"
    End With
    Call loadDataToList(Me.lst1) '//lst1 �ڷ� ����
    
    '//ȭ���޺��ڽ�
    setCBox Me.cbo_FX, "FX", Me.Name
    setCBox Me.cbo1, "FX", Me.Name
    setCBox Me.cbo2, "FX", Me.Name
    
    Call control_initialize1
    Call control_initialize2
    
    txt_date.SetFocus
End Sub

'--------------------------------
'  �Է��׸� �ʱ�ȭ
'--------------------------------
Private Sub control_initialize1()
    cbo_FX.ListIndex = -1: txt_krw = Empty: txt_usd = Empty: txt_date = Empty
    lst1.ListIndex = -1
End Sub
Private Sub control_initialize2()
    txtC1 = Empty: txtC2 = Empty: txtC3 = Empty: cbo1.ListIndex = -1: cbo2.ListIndex = -1
End Sub

'-----------------------------------------
'  ����Ʈ Ŭ�� �̺�Ʈ
'-----------------------------------------
Private Sub lst1_Click()
    With Me '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
        .cbo_FX = .lst1.Column(0, .lst1.ListIndex)
        .txt_date = lst1.Column(2, lst1.ListIndex)
        .txt_krw = Format(.lst1.Column(3, .lst1.ListIndex), "#,##0.000")
        .txt_usd = Format(.lst1.Column(4, .lst1.ListIndex), "#,##0.000000")
    End With
End Sub

'--------------
'  ��¥����
'--------------
Private Sub lbl_date_Click()
    txt_date = Format(Date, "YYYY-MM-DD")
End Sub
Private Sub txt_date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txt_date = Empty Then Exit Sub
    If Me.txt_date = 1 Then Me.txt_date = Date
    If checkTextBox(Me.txt_date, "ȯ����ȸ��", True, "DATE", , True) = False Then Exit Sub
    Me.txt_date.Value = Format(Me.txt_date, "YYYY-MM-DD")
End Sub
Private Sub lbl_date2_Click()
    txtC1 = Format(Date, "YYYY-MM-DD")
End Sub
Private Sub txtC1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtC1 = Empty Then Exit Sub
    If Me.txtC1 = 1 Then Me.txtC1 = Date
    If checkTextBox(Me.txtC1, "ȯ����ȸ��", True, "DATE", , True) = False Then Exit Sub
    Me.txtC1.Value = Format(Me.txtC1, "YYYY-MM-DD")
End Sub

'--------------
'  �ݾװ���
'--------------
Private Sub txtC2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If checkTextBox(Me.txtC2, "ȯ����ȸ�ݾ�", True, "NUMERIC", , True) = False Then Exit Sub
    txtC2.Value = Format(txtC2.Value, "#,##0.000")
End Sub

'----------------------------------------------------------------
'  ����Ʈ�ڽ� ������
'    - loadDataToList(����Ʈ�ڽ���, ��ȸid)
'    - ���� �Ǵ� �߰��� �׸��� �ڵ带 queryKey���� �Է�
'----------------------------------------------------------------
Private Sub loadDataToList(argListBox As MSForms.ListBox, Optional ByVal queryKey As String)
    Dim strSQL As String
    Dim listData() As String
    Dim cntRecord As Integer
    Dim i As Integer, j As Integer
    
    Call control_initialize1
    
    '//SQL��
    strSQL = makeSelectSQL

    '//DB���� �ڷ� ȣ���Ͽ� ���ڵ�¿� ��ȯ
    connectTaskDB
    callDBtoRS "loadDataToList", TB1, strSQL, Me.Name
    
    '//���ڵ���� �����͸� listData �迭�� ��ȯ
    If Not rs.EOF Then
        ReDim listData(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB���� ��ȯ�� �迭�� ũ�� ����: ���ڵ���� ���ڵ� ��, �ʵ� ��
        rs.MoveFirst
        For i = 0 To rs.RecordCount - 1
            For j = 0 To rs.Fields.Count - 1
                If IsNull(rs.Fields(j).Value) = True Then
                    listData(i, j) = ""
                Else
                    listData(i, j) = rs.Fields(j).Value
                End If
            Next j
            rs.MoveNext
        Next i
    End If
    disconnectALL
    
    '//�������� ���ڵ� �� ����
    On Error Resume Next
        cntRecord = UBound(listData) - LBound(listData) + 1 '//��ȸ�� ������ ��
    On Error GoTo 0
    If cntRecord = 0 Then
        MsgBox "ȭ�󸮽�Ʈ�� ��ȯ�� DB �����Ͱ� �����ϴ�.", vbInformation, banner
        argListBox.Clear
        Exit Sub
    End If
    
    '//listData �迭�� ��ȯ�� Data�� ����Ʈ�ڽ��� ������
    argListBox.List = listData
    
    '//����Ʈ ��ȸ �Ŀ��� ���� ����, ���� �� �߰��ÿ��� ����/�߰��� �׸����� �̵�
    If queryKey = Empty Then
        argListBox.ListIndex = -1
    Else
        Call returnListPosition(Me, argListBox.Name, CStr(queryKey))
    End If
End Sub

'-----------------------------------------
'  ���Ǻ� Select SQL�� �ۼ�
'    - makeSelectSQL(�˻���, ����)
'    - DB���� ��ȯ�� ����Ʈ �ʵ����
'-----------------------------------------
Private Function makeSelectSQL(Optional ByVal argSTxt As String, Optional ByVal argFTxt As String) As String
    Dim strSQL As String
    '//ȭ��id, ȭ���Ī, ��ȸ��, ��ȭȯ��, �޷�ȭȯ��
    strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                  "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    makeSelectSQL = strSQL
End Function

'-----------------------------------------
'  ������ �߰�
'-----------------------------------------
Private Sub Cmd_add_Click()
    If MsgBox("ȭ�� �߰��ϰڽ��ϱ�?" & Space(7), vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    caseSave = 1
    Call data_save
End Sub
'-----------------------------------------
'  ������ ����
'-----------------------------------------
Private Sub cmd_delete_Click()
    If Me.lst1.ListIndex = -1 Then Exit Sub
    If MsgBox("ȭ�� �����ϰڽ��ϱ�?" & Space(7), vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    caseSave = 4
    Call data_save
    Call control_initialize1
End Sub

'-----------------------------------------
'  ������ ����:�߰�/����/����
'-----------------------------------------
Private Sub data_save()
    Dim argData As t_currency_cal
    Dim strSQL As String
    Dim result As t_result
    Dim dataType As Integer
        
    '//�Է°���
    If cbo_FX.ListIndex = -1 Then
        MsgBox "ȭ�� �����ϼ���.", vbInformation, banner
        Exit Sub
    End If
    
    '//�ߺ��Է� üũ
    If Me.lst1.ListIndex = -1 Then
        dataType = caseSave
    ElseIf Me.lst1.Column(0, Me.lst1.ListIndex) <> cbo_FX.Column(0, cbo_FX.ListIndex) And caseSave = 2 Then '������ ����Ʈ�� �����Ϸ��� �׸��� �ٸ� ��� �űԵ����ͷ� ����
        dataType = 1
    Else
        dataType = caseSave
    End If
    If checkDoubleInput2(dataType, "currency_id", "user_id", Me.cbo_FX.Column(0, cbo_FX.ListIndex), user_id, TB1, Me.Name) = True Then
        MsgBox "������ ȭ����� �����մϴ�. Ȯ���� �ּ���.", vbInformation, banner
        Exit Sub
    End If
    
    '//������ ����ü�� ��ȯ(���Ա����� �°�)
    With argData
        .currency_id = cbo_FX.Column(0, cbo_FX.ListIndex)
        .currency_un = cbo_FX.Column(1, cbo_FX.ListIndex)
        .refer_dt = Date
        .fx_rate_krw = FXRateC(cbo_FX.Column(1, cbo_FX.ListIndex), Date, 0) '//���ó�¥ ȯ��(��ȭ)
        .fx_rate_usd = FXRateC(cbo_FX.Column(1, cbo_FX.ListIndex), Date, 1) '//���ó�¥ ȯ��(�޷�ȭ)
        .user_id = user_id
    End With
    
    '//����Ʈ ��ġ ��ȯ�� ���� id
    If Me.lst1.ListIndex = -1 Then
        queryKey = 0
    Else
        queryKey = Me.lst1.Column(0, Me.lst1.ListIndex)
    End If
    
    '//������ ���� ���̽��� ���� ����
    Select Case caseSave
        Case 1: result = InsertData(argData)
        Case 4: result = PDeleteData(argData)
    End Select
    
    '//�������: �޽��� �ڽ�, �α� ���
    Select Case caseSave
        Case 1
            MsgBox "ȭ�� " & result.affectedCount & "�� �߰��Ǿ����ϴ�.", vbInformation, banner
            writeLog "InsertData", TB1, result.strSQL, 0, Me.Name, "ȭ�� ����", result.affectedCount
        Case 4
            MsgBox "ȭ�� " & result.affectedCount & "�� (����)�����Ǿ����ϴ�.", vbInformation, banner
            writeLog "PDeleteData", TB1, result.strSQL, 0, Me.Name, "ȭ�� ��������", result.affectedCount
    End Select
    
    '//����Ʈ�� �ݿ�
    loadDataToList Me.lst1, queryKey
End Sub

'-----------------------------------------
'  ������ �߰�(Insert)
'-----------------------------------------
Private Function InsertData(ByRef argData As t_currency_cal) As t_result
    Dim strSQL As String
    Dim resultCode As Integer
    
    connectTaskDB
    strSQL = "INSERT INTO " & TB1 & "(currency_id, currency_un, refer_dt, fx_rate_krw, fx_rate_usd, user_id) VALUES(" & _
                  argData.currency_id & ", " & _
                  SText(argData.currency_un) & ", " & _
                  SText(argData.refer_dt) & ", " & _
                  argData.fx_rate_krw & ", " & _
                  argData.fx_rate_usd & ", " & _
                  argData.user_id & ");"
                  
    '//���� �� ��� ��ȯ
    InsertData.affectedCount = executeSQL("InsertData", TB1, strSQL, Me.Name, "ȭ�� �߰�")
    InsertData.strSQL = strSQL
    
    '//������ �Է� id ��ȯ
    queryKey = cbo_FX.Column(0, cbo_FX.ListIndex)
    
    disconnectALL
End Function

'-----------------------------------------------------------
'  ������ ��������(Delete)
'-----------------------------------------------------------
Private Function PDeleteData(ByRef argData As t_currency_cal) As t_result
    Dim strSQL As String
    Dim cntData As Integer
    Dim affectedCount As Long
    
    '//������ ����
    connectTaskDB
    strSQL = "DELETE FROM " & TB1 & " " & _
                  " WHERE currency_id = " & argData.currency_id & ";"
    
    '//���� �� ��� ��ȯ
    PDeleteData.affectedCount = executeSQL("PDeleteData", TB1, strSQL, Me.Name, "ȭ�� ��������")
    PDeleteData.strSQL = strSQL
    
    disconnectALL
End Function

'-----------------------------------------
'  ������ ���� �ۼ�
'-----------------------------------------
Private Sub cmd_new_Click() '�����ۼ�
    Call control_initialize1
    lst1.ListIndex = -1
End Sub
Private Sub cmd_Cnew_Click()
    control_initialize2
End Sub

'-----------------------------------------
'  �� �ݱ�
'-----------------------------------------
Private Sub cmd_close_Click()
    Unload Me
End Sub

'----------------------
'  ȯ��������Ʈ
'----------------------
Private Sub cmd_update_Click()
    Dim strSQL As String
    Dim strCurrencyID As String
    Dim i As Integer, k As Long
    Dim queryDate As Variant
    
    '//��ȸ��
    If txt_date = Empty Then txt_date.Value = Date
    queryDate = txt_date.Value
    
    '//��ȸ��� ȭ��
    If lst1.ListCount = 0 Then
        MsgBox "��ϵ� ȭ�� �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
    '//ȯ�� ������Ʈ
    Call connectTaskDB
    For i = 0 To Me.lst1.ListCount - 1
        strCurrencyID = Me.lst1.Column(0, i)
        strSQL = "UPDATE " & TB1 & " " & _
                      "SET refer_dt = " & SText(CDate(queryDate)) & ", " & _
                            "fx_rate_krw = " & FXRateC(Me.lst1.Column(1, i), CDate(queryDate), 0) & ", " & _
                            "fx_rate_usd = " & FXRateC(Me.lst1.Column(1, i), CDate(queryDate), 1) & " " & _
                      "WHERE currency_id = " & SText(strCurrencyID) & " AND user_id = " & user_id & ";"
        k = k + executeSQL("cmd_update_Click", TB1, strSQL, Me.Name, "ȯ�� ������Ʈ")
    Next i
    Call disconnectALL
    loadDataToList Me.lst1
    
    MsgBox "������ ���� ȯ���� ������Ʈ �Ǿ����ϴ�." & Space(7) & vbNewLine & vbNewLine & _
                  "����Ʈ ȭ�� �� : " & Me.lst1.ListCount & "��" & vbNewLine & _
                  "������Ʈ�� ȭ�� �� : " & k & "��", vbInformation, banner
    writeLog "cmd_update_Click", TB1, strSQL, 0, Me.Name, "ȯ�� ������Ʈ", k
                  
    Me.lst1.ListIndex = 0
End Sub

'----------------------
'  ȯ�� ��ȸ
'----------------------
Private Sub cmd_refer_Click()
    '//�Է� ����
    If txtC1 = Empty Then
        MsgBox "ȯ����ȸ���� �Է��ϼ���.", vbInformation, banner
        txtC1.SetFocus
        Exit Sub
    End If
    If cbo1.ListIndex = -1 Then
        MsgBox "��ȸȰ ȭ�� �����ϼ���.", vbInformation, banner
        cbo1.SetFocus
        Exit Sub
    End If
    If cbo2.ListIndex = -1 Then
        MsgBox "��ȸȰ ȭ�� �����ϼ���.", vbInformation, banner
        cbo2.SetFocus
        Exit Sub
    End If
    If txtC2 = Empty Then
        MsgBox "��ȸ�� �ݾ��� �Է��ϼ���.", vbInformation, banner
        txtC2.SetFocus
        Exit Sub
    End If
    
    '//ȯ�����
    txtC3 = Empty
    If cbo1.Column(1, cbo1.ListIndex) <> "KRW" And cbo2.Column(1, cbo2.ListIndex) = "KRW" Then
        txtC3 = Format(FXRateC(cbo1.Column(1, cbo1.ListIndex), CDate(txtC1), 0) * txtC2, "#,##0.0000")
        
    ElseIf cbo1.Column(1, cbo1.ListIndex) <> "KRW" And cbo1.Column(1, cbo1.ListIndex) <> "USD" And cbo2.Column(1, cbo2.ListIndex) = "USD" Then
        txtC3 = Format(FXRateC(cbo1.Column(1, cbo1.ListIndex), CDate(txtC1), 1) * txtC2, "#,##0.0000")
        
    ElseIf cbo1.Column(1, cbo1.ListIndex) = "KRW" And cbo2.Column(1, cbo2.ListIndex) <> "KRW" Then
        txtC3 = Format((1 / FXRateC(cbo2.Column(1, cbo2.ListIndex), CDate(txtC1), 0)) * txtC2, "#,##0.0000")
        
    ElseIf cbo1.Column(1, cbo1.ListIndex) = "USD" And cbo2.Column(1, cbo2.ListIndex) <> "USD" Then
        txtC3 = Format((1 / FXRateC(cbo2.Column(1, cbo2.ListIndex), CDate(txtC1), 1)) * txtC2, "#,##0.0000")
        
    ElseIf cbo1.Column(1, cbo1.ListIndex) <> "KRW" And cbo1.Column(1, cbo1.ListIndex) <> "USD" And cbo2.Column(1, cbo2.ListIndex) <> "KRW" And cbo2.Column(1, cbo2.ListIndex) <> "USD" Then
        txtC3 = Format((FXRateC(cbo1.Column(1, cbo1.ListIndex), CDate(txtC1), 1) / FXRateC(cbo2.Column(1, cbo2.ListIndex), CDate(txtC1), 1)) * txtC2, "#,##0.0000")
        
    End If
    
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_unit_converter 
   Caption         =   "단위변환기"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6120
   OleObjectBlob   =   "f_unit_converter.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "f_unit_converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------
'  폼 종료 시
'--------------
Private Sub UserForm_Terminate()

End Sub

'-------------------------------
'  단위변환 폼(common)
'-------------------------------
Private Sub UserForm_Initialize()
    Dim strSQL As String

    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    
    '//기초설정
'    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    
    '//텍스트박스
    txt1.Value = 1
    
    '//단위 종류 콤보박스
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
'  단위 콤보박스 구성
'-----------------------
Private Sub make_cbo_unit1()
    Dim strSQL As String
    '//단위1 콤보박스
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
    '//단위2 콤보박스
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
'  카테고리 바뀌면
'    - 단위1 변경
'    - 단위2 변경
'----------------------
Private Sub cbo_unit_Change()
    Me.txt1.Value = 1
    Call make_cbo_unit1
    Call make_cbo_unit2
    txt1.SetFocus
End Sub

'----------------------
'  단위변환
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
    Dim unit1S As Boolean, unit2S As Boolean '변환하려는 단위가 기본단위인지 확인
    Dim amt1 As Double, amt2 As Double '변화하려는 단위 2개 모두 기본단위가 아닌 경우 연산을 위해 / currency로 데이터형식을 지정할 경우 소수점 4자리밖에 안되서 double로 지정
    Dim strSQL As String
    
    '//입력검증: checkTextBox(텍스트박스 이름, 텍스트박스 타이틀, 필수여부, 데이터형, 길이 제한, 입력 후 포커싱)
    If Me.txt1 = Empty Then
        MsgBox "단위를 변환할 수치를 입력하세요.", vbInformation, Banner
        Exit Sub
    Else
        If checkTextBox(Me.txt1, "단위변환할 수치", True, "NUMERIC", , True) = False Then Exit Sub
    End If
    If cbo_unit1.ListIndex = -1 Then
        MsgBox "변환할 단위를 입력하세요.", vbInformation, Banner
        Exit Sub
    End If
    If cbo_unit2.ListIndex = -1 Then
        MsgBox "변환할 단위를 입력하세요.", vbInformation, Banner
        Exit Sub
    End If
    
    '//단위변환
        '기본단위 여부 확인
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
        '케이스별 단위변환
        If (cbo_unit1.Column(0, cbo_unit1.ListIndex) = cbo_unit2.Column(0, cbo_unit2.ListIndex)) Then '두 단위가 동일한 경우
            amt1 = 1
        ElseIf unit1S = True And unit2S = False Then '단위1은 기본단위, 단위2는 아닌 경우
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & " AND a.unit_id2 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "단위변환값조회"
            If rs.EOF <> True Then
                amt1 = rs("value").Value
            End If
            disconnectALL
        ElseIf unit1S = False And unit2S = True Then '단위2만 기본단위인 경우
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id2 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & " AND a.unit_id1 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "단위변환값조회"
            If rs.EOF <> True Then
                amt1 = 1 / rs("value").Value
            End If
            disconnectALL
        ElseIf unit1S = False And unit2S = False Then '둘 다 기본단위가 아닌 경우
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit1.Column(2, cbo_unit1.ListIndex) & " AND a.unit_id2 = " & cbo_unit1.Column(0, cbo_unit1.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "단위변환값조회"
            If rs.EOF <> True Then
                amt1 = 1 / rs("value").Value
            End If
            disconnectALL
            strSQL = "SELECT a.value FROM co_account.unit_conversion a WHERE a.unit_id1 = " & cbo_unit2.Column(2, cbo_unit2.ListIndex) & " AND a.unit_id2 = " & cbo_unit2.Column(0, cbo_unit2.ListIndex) & ";"
            connectTaskDB
            callDBtoRS "cmd_refer_Click", "co_account.unit_conversion", strSQL, Me.Name, "단위변환값조회"
            If rs.EOF <> True Then
                amt2 = rs("value").Value
            End If
            disconnectALL
            amt1 = amt1 * amt2
        End If
    Call adjust_decimal(Me.txt2, Round(txt1.Value * amt1, 10)) 'decimal을 10자리까지로 짤라서 진행
End Sub

'----------------------------
'  소소점 자리수 동적 변환
'----------------------------
Private Sub adjust_decimal(argTB As MSForms.TextBox, argValue As Double)
    Dim noA As Integer '전체 글자수
    Dim noB As Integer '마지막 '0'값의 위치
    Dim noC As Integer '소수점 자리수
    
    noA = Len(Format(argValue, "@")) '숫자의 경우 len함수가 안먹어서 문자형으로 변환하여 연산
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
'  폼 닫기
'-----------------------------------------
Private Sub cmd_close_Click()
    Unload Me
End Sub

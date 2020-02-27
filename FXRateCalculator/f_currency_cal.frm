VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_currency_cal 
   Caption         =   "환율조회기"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8640
   OleObjectBlob   =   "f_currency_cal.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "f_currency_cal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cntlst1Col As Integer '//리스트 컬럼 수
Dim TB1 As String '//폼에 연결된 DB 테이블
Const strOrderBy As String = "currency_un ASC" '//DB에서 sort_order 필드
Dim caseSave As Integer '//1: 추가, 2: 수정, 3: 삭제(SUSPEND), 4: 완전삭제
Dim queryKey As Integer '//리스트 위치 반환에 사용될 id

'--------------
'  폼 종료 시
'--------------
Private Sub UserForm_Terminate()

End Sub

'-------------------------------
'  환율계산 폼(common)
'-------------------------------
Private Sub UserForm_Initialize()
    If connIP = Empty Then setGlobalVariant '//전역변수 재설정
    If checkLogin = 0 Then f_login.Show '//로그인체크
    
    '//기초설정
    cntlst1Col = 5 '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율, 정렬순서
'    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    
    '//폼에 연결된 object 정보
    TB1 = "fx_calculator.currency_cal"
    
    With txtC3
        .Locked = True
        .BackColor = &H80000018
    End With
        
    '//리스트상자 설정
    With lst1
        .ColumnCount = cntlst1Col
        .ColumnHeads = False
        .ColumnWidths = "0,48,70,70,70" '화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        .Width = 260
        .TextAlign = fmTextAlignLeft
        .Font = "맑은 고딕"
    End With
    Call loadDataToList(Me.lst1) '//lst1 자료 구성
    
    '//화폐콤보박스
    setCBox Me.cbo_FX, "FX", Me.Name
    setCBox Me.cbo1, "FX", Me.Name
    setCBox Me.cbo2, "FX", Me.Name
    
    Call control_initialize1
    Call control_initialize2
    
    txt_date.SetFocus
End Sub

'--------------------------------
'  입력항목 초기화
'--------------------------------
Private Sub control_initialize1()
    cbo_FX.ListIndex = -1: txt_krw = Empty: txt_usd = Empty: txt_date = Empty
    lst1.ListIndex = -1
End Sub
Private Sub control_initialize2()
    txtC1 = Empty: txtC2 = Empty: txtC3 = Empty: cbo1.ListIndex = -1: cbo2.ListIndex = -1
End Sub

'-----------------------------------------
'  리스트 클릭 이벤트
'-----------------------------------------
Private Sub lst1_Click()
    With Me '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
        .cbo_FX = .lst1.Column(0, .lst1.ListIndex)
        .txt_date = lst1.Column(2, lst1.ListIndex)
        .txt_krw = Format(.lst1.Column(3, .lst1.ListIndex), "#,##0.000")
        .txt_usd = Format(.lst1.Column(4, .lst1.ListIndex), "#,##0.000000")
    End With
End Sub

'--------------
'  날짜관련
'--------------
Private Sub lbl_date_Click()
    txt_date = Format(Date, "YYYY-MM-DD")
End Sub
Private Sub txt_date_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txt_date = Empty Then Exit Sub
    If Me.txt_date = 1 Then Me.txt_date = Date
    If checkTextBox(Me.txt_date, "환율조회일", True, "DATE", , True) = False Then Exit Sub
    Me.txt_date.Value = Format(Me.txt_date, "YYYY-MM-DD")
End Sub
Private Sub lbl_date2_Click()
    txtC1 = Format(Date, "YYYY-MM-DD")
End Sub
Private Sub txtC1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtC1 = Empty Then Exit Sub
    If Me.txtC1 = 1 Then Me.txtC1 = Date
    If checkTextBox(Me.txtC1, "환율조회일", True, "DATE", , True) = False Then Exit Sub
    Me.txtC1.Value = Format(Me.txtC1, "YYYY-MM-DD")
End Sub

'--------------
'  금액관련
'--------------
Private Sub txtC2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If checkTextBox(Me.txtC2, "환율조회금액", True, "NUMERIC", , True) = False Then Exit Sub
    txtC2.Value = Format(txtC2.Value, "#,##0.000")
End Sub

'----------------------------------------------------------------
'  리스트박스 리스팅
'    - loadDataToList(리스트박스명, 조회id)
'    - 수정 또는 추가된 항목의 코드를 queryKey값에 입력
'----------------------------------------------------------------
Private Sub loadDataToList(argListBox As MSForms.ListBox, Optional ByVal queryKey As String)
    Dim strSQL As String
    Dim listData() As String
    Dim cntRecord As Integer
    Dim i As Integer, j As Integer
    
    Call control_initialize1
    
    '//SQL문
    strSQL = makeSelectSQL

    '//DB에서 자료 호출하여 레코드셋에 반환
    connectTaskDB
    callDBtoRS "loadDataToList", TB1, strSQL, Me.Name
    
    '//레코드셋의 데이터를 listData 배열에 반환
    If Not rs.EOF Then
        ReDim listData(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB에서 반환할 배열의 크기 지정: 레코드셋의 레코드 수, 필드 수
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
    
    '//리스팅할 레코드 수 검토
    On Error Resume Next
        cntRecord = UBound(listData) - LBound(listData) + 1 '//조회된 데이터 수
    On Error GoTo 0
    If cntRecord = 0 Then
        MsgBox "화폐리스트에 반환할 DB 데이터가 없습니다.", vbInformation, banner
        argListBox.Clear
        Exit Sub
    End If
    
    '//listData 배열로 반환된 Data를 리스트박스에 리스팅
    argListBox.List = listData
    
    '//리스트 조회 후에는 선택 없음, 수정 및 추가시에서 수정/추가된 항목으로 이동
    If queryKey = Empty Then
        argListBox.ListIndex = -1
    Else
        Call returnListPosition(Me, argListBox.Name, CStr(queryKey))
    End If
End Sub

'-----------------------------------------
'  조건별 Select SQL문 작성
'    - makeSelectSQL(검색어, 필터)
'    - DB에서 반환할 리스트 필드수정
'-----------------------------------------
Private Function makeSelectSQL(Optional ByVal argSTxt As String, Optional ByVal argFTxt As String) As String
    Dim strSQL As String
    '//화폐id, 화폐약칭, 조회일, 원화환율, 달러화환율
    strSQL = "SELECT a.currency_id, a.currency_un, a.refer_dt, a.fx_rate_krw, a.fx_rate_usd " & _
                  "FROM " & TB1 & " a WHERE a.user_id = " & user_id & ";"
    makeSelectSQL = strSQL
End Function

'-----------------------------------------
'  데이터 추가
'-----------------------------------------
Private Sub Cmd_add_Click()
    If MsgBox("화폐를 추가하겠습니까?" & Space(7), vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    caseSave = 1
    Call data_save
End Sub
'-----------------------------------------
'  데이터 삭제
'-----------------------------------------
Private Sub cmd_delete_Click()
    If Me.lst1.ListIndex = -1 Then Exit Sub
    If MsgBox("화폐를 삭제하겠습니까?" & Space(7), vbQuestion + vbYesNo, banner) = vbNo Then
        Exit Sub
    End If
    caseSave = 4
    Call data_save
    Call control_initialize1
End Sub

'-----------------------------------------
'  데이터 저장:추가/수정/삭제
'-----------------------------------------
Private Sub data_save()
    Dim argData As t_currency_cal
    Dim strSQL As String
    Dim result As t_result
    Dim dataType As Integer
        
    '//입력검증
    If cbo_FX.ListIndex = -1 Then
        MsgBox "화폐를 선택하세요.", vbInformation, banner
        Exit Sub
    End If
    
    '//중복입력 체크
    If Me.lst1.ListIndex = -1 Then
        dataType = caseSave
    ElseIf Me.lst1.Column(0, Me.lst1.ListIndex) <> cbo_FX.Column(0, cbo_FX.ListIndex) And caseSave = 2 Then '선택한 리스트와 수정하려는 항목이 다를 경우 신규데이터로 검증
        dataType = 1
    Else
        dataType = caseSave
    End If
    If checkDoubleInput2(dataType, "currency_id", "user_id", Me.cbo_FX.Column(0, cbo_FX.ListIndex), user_id, TB1, Me.Name) = True Then
        MsgBox "동일한 화폐명이 존재합니다. 확인해 주세요.", vbInformation, banner
        Exit Sub
    End If
    
    '//데이터 구조체로 반환(테입구조에 맞게)
    With argData
        .currency_id = cbo_FX.Column(0, cbo_FX.ListIndex)
        .currency_un = cbo_FX.Column(1, cbo_FX.ListIndex)
        .refer_dt = Date
        .fx_rate_krw = FXRateC(cbo_FX.Column(1, cbo_FX.ListIndex), Date, 0) '//오늘날짜 환율(원화)
        .fx_rate_usd = FXRateC(cbo_FX.Column(1, cbo_FX.ListIndex), Date, 1) '//오늘날짜 환율(달러화)
        .user_id = user_id
    End With
    
    '//리스트 위치 반환에 사용될 id
    If Me.lst1.ListIndex = -1 Then
        queryKey = 0
    Else
        queryKey = Me.lst1.Column(0, Me.lst1.ListIndex)
    End If
    
    '//데이터 저장 케이스에 따라 실행
    Select Case caseSave
        Case 1: result = InsertData(argData)
        Case 4: result = PDeleteData(argData)
    End Select
    
    '//결과보고: 메시지 박스, 로그 기록
    Select Case caseSave
        Case 1
            MsgBox "화폐가 " & result.affectedCount & "건 추가되었습니다.", vbInformation, banner
            writeLog "InsertData", TB1, result.strSQL, 0, Me.Name, "화폐 저장", result.affectedCount
        Case 4
            MsgBox "화폐가 " & result.affectedCount & "건 (완전)삭제되었습니다.", vbInformation, banner
            writeLog "PDeleteData", TB1, result.strSQL, 0, Me.Name, "화폐 완전삭제", result.affectedCount
    End Select
    
    '//리스트에 반영
    loadDataToList Me.lst1, queryKey
End Sub

'-----------------------------------------
'  데이터 추가(Insert)
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
                  
    '//실행 및 결과 반환
    InsertData.affectedCount = executeSQL("InsertData", TB1, strSQL, Me.Name, "화폐 추가")
    InsertData.strSQL = strSQL
    
    '//마지막 입력 id 반환
    queryKey = cbo_FX.Column(0, cbo_FX.ListIndex)
    
    disconnectALL
End Function

'-----------------------------------------------------------
'  데이터 완전삭제(Delete)
'-----------------------------------------------------------
Private Function PDeleteData(ByRef argData As t_currency_cal) As t_result
    Dim strSQL As String
    Dim cntData As Integer
    Dim affectedCount As Long
    
    '//데이터 삭제
    connectTaskDB
    strSQL = "DELETE FROM " & TB1 & " " & _
                  " WHERE currency_id = " & argData.currency_id & ";"
    
    '//실행 및 결과 반환
    PDeleteData.affectedCount = executeSQL("PDeleteData", TB1, strSQL, Me.Name, "화폐 완전삭제")
    PDeleteData.strSQL = strSQL
    
    disconnectALL
End Function

'-----------------------------------------
'  데이터 새로 작성
'-----------------------------------------
Private Sub cmd_new_Click() '새로작성
    Call control_initialize1
    lst1.ListIndex = -1
End Sub
Private Sub cmd_Cnew_Click()
    control_initialize2
End Sub

'-----------------------------------------
'  폼 닫기
'-----------------------------------------
Private Sub cmd_close_Click()
    Unload Me
End Sub

'----------------------
'  환율업데이트
'----------------------
Private Sub cmd_update_Click()
    Dim strSQL As String
    Dim strCurrencyID As String
    Dim i As Integer, k As Long
    Dim queryDate As Variant
    
    '//조회일
    If txt_date = Empty Then txt_date.Value = Date
    queryDate = txt_date.Value
    
    '//조회대상 화폐
    If lst1.ListCount = 0 Then
        MsgBox "등록된 화폐가 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
    '//환율 업데이트
    Call connectTaskDB
    For i = 0 To Me.lst1.ListCount - 1
        strCurrencyID = Me.lst1.Column(0, i)
        strSQL = "UPDATE " & TB1 & " " & _
                      "SET refer_dt = " & SText(CDate(queryDate)) & ", " & _
                            "fx_rate_krw = " & FXRateC(Me.lst1.Column(1, i), CDate(queryDate), 0) & ", " & _
                            "fx_rate_usd = " & FXRateC(Me.lst1.Column(1, i), CDate(queryDate), 1) & " " & _
                      "WHERE currency_id = " & SText(strCurrencyID) & " AND user_id = " & user_id & ";"
        k = k + executeSQL("cmd_update_Click", TB1, strSQL, Me.Name, "환율 업데이트")
    Next i
    Call disconnectALL
    loadDataToList Me.lst1
    
    MsgBox "다음과 같이 환율이 업데이트 되었습니다." & Space(7) & vbNewLine & vbNewLine & _
                  "리스트 화폐 수 : " & Me.lst1.ListCount & "개" & vbNewLine & _
                  "업데이트된 화폐 수 : " & k & "개", vbInformation, banner
    writeLog "cmd_update_Click", TB1, strSQL, 0, Me.Name, "환율 업데이트", k
                  
    Me.lst1.ListIndex = 0
End Sub

'----------------------
'  환율 조회
'----------------------
Private Sub cmd_refer_Click()
    '//입력 검증
    If txtC1 = Empty Then
        MsgBox "환율조회일을 입력하세요.", vbInformation, banner
        txtC1.SetFocus
        Exit Sub
    End If
    If cbo1.ListIndex = -1 Then
        MsgBox "조회활 화폐를 선택하세요.", vbInformation, banner
        cbo1.SetFocus
        Exit Sub
    End If
    If cbo2.ListIndex = -1 Then
        MsgBox "조회활 화폐를 선택하세요.", vbInformation, banner
        cbo2.SetFocus
        Exit Sub
    End If
    If txtC2 = Empty Then
        MsgBox "조회할 금액을 입력하세요.", vbInformation, banner
        txtC2.SetFocus
        Exit Sub
    End If
    
    '//환율계산
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


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_login 
   Caption         =   "로그인"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   OleObjectBlob   =   "f_login.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "f_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------
'  로그인창 종료 시 로그인검증
'-----------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    If checkLogin = 0 Then
        MsgBox "로그인 정보가 확인되지 않았습니다." & Space(7) & vbNewLine & _
            "프로그램을 종료합니다.", vbInformation, Banner
        ThisWorkbook.Close savechanges:=False
    End If
End Sub

'------------------------------------------------------
'  로그인 폼(common)
'  - 프로그램 버전 체크하여 과거 버전 사용 제한
'  - IP기록
'------------------------------------------------------
Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim strSQL As String
    
    '//기초설정
    Me.cmd_close.Width = 0
    Me.cmd_close.Cancel = True
    Me.Caption = Banner
    txt1.Value = Application.UserName
    Me.lbl_pv = programv
    Me.lbl_report = reportfile_nm
        
    '//등록된 사용자 체크
    If checkUserNm(txt1.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & Space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경해 주세요." & Space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청해 주세요.", vbInformation, Banner
        GoTo n
    End If
    
    '//비밀번호 설정 여부 체크
    Call checkInitialPW
n:
    txt2.SetFocus
    Exit Sub
ErrHandler:
    End
End Sub

'-------------------------------------------------------------------------------------
'  등록된 사용자 체크
'    - txt1에 입력된 사용자가 등록된 사용자인지 검토하여 true / false 값 반환
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
'  txt1에서 exit 할 경우
'    - 사용자 이름 등록여부 체크
'    - 비밀번호 초기 설정 여부 체크
'---------------------------------------
Private Sub txt1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If txt1 = Empty Then
        Exit Sub
    End If
    
    '//사용자 이름 등록 여부 체크
    If checkUserNm(txt1.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & Space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경해 주세요." & Space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청해 주세요.", vbInformation, Banner
        txt1.SetFocus
        Exit Sub
    End If
    If txt1.Value <> Application.UserName Then
        Application.UserName = txt1.Value
    End If
    
    '//비밀번호 초기 설정 여부 체크
    Call checkInitialPW
    
End Sub

'----------------------------------------------------------------------------------------
'  등록된 사용자의 경우 비밀번호가 설정되어 있었는지 체크하고 설정하도록 진행
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
    If strPW = 1 Then '//PW 입력 이력이 없으면 PW 설정
        MsgBox "비밀번호가 설정되어 있지 않습니다.", vbInformation, Banner
        registerNewPW
    End If
    disconnectALL
End Sub

'-----------------------
'  신규비밀번호 등록
'-----------------------
Private Sub registerNewPW()
    Dim strSQL As String
    Dim strPW As Integer
    Dim user_pw As Variant
    Dim affectedCount As Long
    '비밀번호 입력 받기
    Do
        user_pw = InputBoxPW("신규 비밀번호를 대소문자를 구분하여 4자리 이상으로 설정해 주세요.", Banner)
    Loop Until user_pw <> Empty And Len(user_pw) > 3
    '비밀번호 등록
    strSQL = "UPDATE common.users SET user_pw = SHA2(" & SText(user_pw) & ", 512) WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    affectedCount = executeSQL("checkInitialPW", "common.users", strSQL, "f_login", "초기비밀번호설정")
    '비밀번호 초기화 비활성화
    strSQL = "UPDATE common.users SET pw_initialize = 0 WHERE user_id = (SELECT user_id FROM common.users WHERE user_nm = " & SText(Me.txt1.Value) & ");"
    affectedCount = executeSQL("checkInitialPW", "common.users", strSQL, "f_login", "초기비밀번호설정")
    If affectedCount = 0 Then
        MsgBox "비밀번호가 설정되지 않았습니다." & Space(7) & vbNewLine & _
            "관리자에게 문의하여 주시기 바랍니다.", vbInformation, Banner
        'ThisWorkbook.Close savechanges:=False
    Else
        MsgBox "비밀번호 설정이 완료되었습니다." & Space(7), vbInformation, Banner
    End If
    disconnectALL
End Sub

'---------------------------------------
'  확인버튼 시
'    - 사용자 이름 등록 여부 체크
'    - 프로그램 최신버전 확인
'    - IP체크
'    - 비밀번호 맞는 지 검토
'    - 법인회계런처 단축키 설정: ALT + 왼쪽방향키
'    - 환영인사
'---------------------------------------
Private Sub cmd_query_Click()
    Dim strSQL As String
    Dim affectedCount As Long
    Dim ipRng As Integer
    
    '//사용자 이름 등록 여부 체크
    If txt1 = Empty Then
        MsgBox "사용자의 이름을 입력하세요.", vbInformation, Banner
        Exit Sub
    End If
    If checkUserNm(txt1.Value) = False Then
        MsgBox "등록된 사용자가 아닙니다." & Space(7) & vbNewLine & _
                "로그인 창에서 이름을 변경해 주세요." & Space(7) & vbNewLine & _
                "사용자 등록이 필요하면 관리자에게 요청해 주세요.", vbInformation, Banner
        txt1.SetFocus
        Exit Sub
    End If
    If txt1.Value <> Application.UserName Then
        Application.UserName = txt1.Value
    End If
    
    '//비밀번호 입력 여부 체크
    If txt2 = Empty Then
        MsgBox "비밀번호를 입력하세요.", vbInformation, Banner
        txt2.SetFocus
        Exit Sub
    End If
    
    '//프로그램 버전 확인
    strSQL = "SELECT programv FROM common.users WHERE user_id = 3"
    connectCommonDB
    callDBtoRS "txt1_Exit", "common.users", strSQL, Me.Name, "프로그램버전 확인"
    If rs("programv").Value <> programv Then
        MsgBox "사용하려는 법인회계프로그램이 최신버전이 아닙니다." & vbNewLine & _
                     "프로그램 오류 방지를 위해 최신버전으로 사용해 주세요.", vbInformation, Banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//IP확인
    ipRng = Mid(GetLocalIPaddress, InStr(5, GetLocalIPaddress, ".") + 2, 2)
    If ipRng <> 10 And ipRng <> 11 Then
        MsgBox "법인회계프로그램은 허용된 PC에서만 사용 가능합니다." & vbNewLine & _
                     "프로그램을 종료합니다.", vbInformation, Banner
        disconnectALL
        cmd_close_Click
    End If
    
    '//비밀번호 맞는 지 검토
    If checkPW(txt2.Value) = True Then
        '비밀번호가 맞으면 Welcome
        checkLogin = 1
        setGlobalVariant
        '//법인회계 런처 단축키 설정
        Application.OnKey "%{LEFT}", "start_co_account"
        '//DB에 IP입력
        '[기존IP지우기]
        strSQL = "UPDATE common.users SET user_ip = NULL WHERE user_id = " & user_id & ";"
        connectCommonDB
        affectedCount = executeSQL("cmd_query_Click", "common.users", strSQL, Me.Name, "사용자IP기록")
        '[신규IP넣기]
        strSQL = "UPDATE common.users SET user_ip = " & SText(GetLocalIPaddress) & " WHERE user_id = " & user_id & ";"
        affectedCount = executeSQL("cmd_query_Click", "common.users", strSQL, Me.Name, "사용자IP기록")
        If affectedCount > 0 Then
            writeLog "cmd_query_Click", "common.users", strSQL, 0, Me.Name, "사용자IP기록", affectedCount
        End If
        '//환영인사
        MsgBox Application.UserName & "님 복많이 받으세요." & Space(7) & vbNewLine & vbNewLine & _
                 "오늘은 " & Format(Date, "YYYY-MM-DD") & "일 입니다." & vbNewLine & _
                 "오늘도 ANIMO!", vbInformation, Banner
        today = Date
        Unload Me
    Else
        '비밀번호가 다르면 다시 입력
        MsgBox "비밀번호가 틀렸습니다." & Space(7) & vbNewLine & _
            "비밀번호를 다시 입력하여 주세요.", vbInformation, Banner
        txt2.Value = Empty
        txt2.SetFocus
        Exit Sub
    End If
        
End Sub

'------------------------------------------------------------------------
'  입력된 비밀번호가 맞는지 틀린지 검토하여 true / false 값 반환
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
'  비밀번호 변경
'    - 기존 비밀번호 확인
'    - 신규 비밀번호 입력
'---------------------------------------
Private Sub cmd_chgPW_Click()
    Dim oldPW As String
    Dim newPW As String
    If MsgBox("비밀번호를 변경하겠습니까?", vbQuestion + vbYesNo, Banner) = vbNo Then
        Exit Sub
    Else
        oldPW = InputBoxPW("기존 비밀번호를 입력하세요.", Banner)
        If StrPtr(oldPW) = 0 Then
            MsgBox "기존 비밀번호 입력이 취소되었습니다.", vbInformation, Banner
            Exit Sub
        Else
            If checkPW(oldPW) = True Then
                registerNewPW
            Else
                MsgBox "기존 비밀번호가 일치하지 않습니다." & vbNewLine & _
                             "관리자에게 문의하여 주세요.", vbInformation, Banner
            End If
        End If
    End If
End Sub


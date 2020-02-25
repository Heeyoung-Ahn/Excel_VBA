Attribute VB_Name = "a_Ribbon"
Option Explicit

'--------------------
'  추가기능 메뉴바
'--------------------
Sub make_menubar()
Call reset_menubar
On Error Resume Next

    With Application.CommandBars("tools").Controls
        With .Add(Type:=msoControlButton)
            .FaceId = 1907
            .Caption = "로그인"
            .OnAction = "LogIn"
        End With
        With .Add(Type:=msoControlButton)
            .FaceId = 5955
            .Caption = "로그아웃"
            .OnAction = "LogOut"
        End With
        With .Add(Type:=msoControlButton)
            .FaceId = 1088
            .Caption = "프로그램종료"
            .OnAction = "AddinUninstall"
        End With
    End With

On Error GoTo 0
End Sub

'---------------------------------------------------------------------
'  로그인
'    - 추가기능이 2개 이상일 경우 프로시저 명을 다르게 해야 함
'---------------------------------------------------------------------
Sub LogIn()
    f_login.Show
End Sub

'------------
'  로그아웃
'------------
Sub LogOut()
    If checkLogin = 0 Then
        MsgBox Application.UserName & "님 이미 로그아웃 되어 있습니다.", vbInformation, banner
        Exit Sub
    End If
    checkLogin = 0 '로그아웃 상태
    '//전역변수 초기화
    connIP = Empty
    connDB = Empty
    connUN = Empty
    connPW = Empty
    user_id = Empty
    user_gb = Empty
    MsgBox "로그아웃 되었습니다." & Space(7), vbInformation, banner
End Sub

'--------------------------
'  추가기능 메뉴바 제거
'--------------------------
Sub reset_menubar()
On Error Resume Next
    Application.CommandBars("WorkSheet Menu Bar").Reset
On Error GoTo 0
End Sub

'------------------
'  추가기능 종료
'------------------
Sub AddinUninstall()
    reset_menubar
    ThisWorkbook.Close False
End Sub


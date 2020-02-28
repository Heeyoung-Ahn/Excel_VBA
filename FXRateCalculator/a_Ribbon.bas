Attribute VB_Name = "a_Ribbon"
Option Explicit

'-----------------------------------------------------
'  리본 메뉴의 Button ID에 대한 처리 프로시저
'-----------------------------------------------------
Sub run_RibbonControl(Button As Office.IRibbonControl)
    Select Case Button.ID
        '//프로그램
        Case "FX_Calculator":     Call FX_Calculator
        Case "InsertPicture":    Call InsertPicture
        Case "InsertPicture2":    Call InsertPicture2
                      
        '//공통
        Case "LogIn":     Call LogIn
        Case "LogOut":     Call LogOut
        Case "AddinUninstall":     Call AddinUninstall
        
        Case Else:     Call RibbonButton_Error(Button.ID)
    End Select
End Sub

'-------------------------------------------------------------------
'  Button ID에 대한 처리 프로시저가 없는 경우 오류 메시지
'-------------------------------------------------------------------
Sub RibbonButton_Error(sbID As String)
   MsgBox "선택하신 메뉴(" & sbID & ")는 아직 준비가 되어 있지 않습니다.", vbCritical, banner
End Sub

'-----------
'  로그인
'-----------
Sub LogIn()
    If checkLogin = 1 Then
        MsgBox Application.UserName & "님 이미 로그인 되어 있습니다.", vbInformation, banner
        Exit Sub
    End If
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
    user_dept = Empty
    MsgBox "로그아웃 되었습니다." & Space(7), vbInformation, banner
End Sub

'------------------------------------
'  현재 파일을 닫아 리본 탭 닫음
'------------------------------------
Sub AddinUninstall()
   ThisWorkbook.Close False
End Sub

'---------------
'  환율조회기
'---------------
Sub FX_Calculator()
    f_currency_cal.Show vbModeless
End Sub


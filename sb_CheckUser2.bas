Attribute VB_Name = "sb_CheckUser"
Option Explicit
Public Const Banner As String = "등록된 사용자 점검"
Sub OpenEvent()

'#############################################################
'워크북을 실행할 때 등록된 사용자인지 여부를 점검하는 프로시저
'#############################################################

Dim aryUser() As Variant
    
    '--//사용 가능한 사용자 목록을 정의하세요.
    aryUser = Array("최우석", "안지혜", "김정미", "이가희", "안희영") '--//★★ 설정값
    
    '--//사용자 이름을 설정합니다.
    Call sbSetUserName(Application.UserName)
    
    '--//설정한 사용자 이름으 등록된 사용자인지 여부를 점검합니다.
    If fnCheckUserName(Application.UserName, aryUser) = False Then
        MsgBox "등록되지 않은 사용자 입니다. 프로그램을 종료합니다." & vbNewLine & _
                "권한을 요청하려면 담당자에게 문의하세요.", vbCritical, Banner
        ThisWorkbook.Close False
    Else
        MsgBox "등록된 사용자 입니다. 반갑습니다.", vbInformation, Banner
    End If

    

End Sub

Sub sbSetUserName(UserNM As String)

'###############################
'사용자 이름을 설정하는 프로시저
'###############################

    '--//현재 사용자 이름을 띄워주고 계속 사용할지 여부를 묻는다.
    If MsgBox("현재 사용자 이름은 " & UserNM & " 입니다." & vbNewLine & _
                "해당 이름을 계속 사용하시겠습니까?", vbQuestion + vbYesNo, Banner) = vbNo Then
Back:
        '--//현재 사용자 이름을 바꾸고 싶을 경우 원하는 값을 입력한다.
        UserNM = InputBox("사용할 사용자 이름을 입력하세요.", Banner, UserNM)
       
       If UserNM = vbNullString Then GoTo Back
       
       '--//입력한 값을 사용할지 다시 한 번 물어본 후 사용자 이름을 해당 값으로 변환한다.
        If MsgBox("사용자 이름을 " & UserNM & "으로 설정 하시겠습니까?", vbQuestion + vbYesNo, Banner) = vbYes Then
            Application.UserName = UserNM
        End If
    End If
    
End Sub

'##########################################################################################
'▶fnCheckUserName 함수
'▶기능: 현재 사용자가 등록된 사용자인지 여부를 점검하여 논리값으로 결과를 반환합니다.
'▶인수설명:
'________UserNM: 설정된 사용자 이름
'________aryUser: 등록된 사용자 이름목록
'##########################################################################################
Function fnCheckUserName(UserNM As String, aryUser As Variant) As Boolean
    Dim var As Variant
    
    '--//초기값 = False(등록되지 않은 사용자)
    fnCheckUserName = False
    
    '--//만약 사용자 이름이 등록된 사용자로 판명될 경우 값을 True로 변경
    For Each var In aryUser
        If var = UserNM Then
            fnCheckUserName = True
        End If
    Next var

End Function

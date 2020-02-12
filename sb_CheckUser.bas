Attribute VB_Name = "sb_CheckUser"
Option Explicit
Public Const banner As String = "엑셀VBA샘플(사용자체크)"
Dim registeredUser As Variant

'------------------------------
'  사용자 이름 지정
'  사용자 이름 등록 확인
'------------------------------
Sub Workbook_Open()
    
    '//프로그램 사용자 설정
    registeredUser = Array("안희영", "사용자2", "사용자3")
    
    '//프로그램 사용자 이름 설정
    Call setUserName(Application.UserName)
    
    '//사용자 이름 등록 여부 확인: 등록되지 않은 사용자는 워크북 종료
    If checkUserName(Application.UserName, registeredUser) = False Then
        MsgBox "'" & Application.UserName & "'님은 프로그램 사용자로 등록되어 있지 않습니다." & vbNewLine & _
            "엑셀 파일을 종료합니다.", vbCritical, banner
        ThisWorkbook.Close savechanges:=False
    Else
        MsgBox "'" & Application.UserName & "'님은 프로그램 사용자로 확인되었습니다." & vbNewLine & _
            "이 파일을 정상적으로 사용할 수 있습니다.", vbInformation, banner
    End If
End Sub

'------------------------
'  사용자 이름 지정
'------------------------
Sub setUserName(userNM As String)
    Dim argURNM As String

    If MsgBox("이 프로그램에서 사용할 사용자의 이름을 " & userNM & "으로 하겠습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        Do
            argURNM = InputBox("담당자의 이름을 입력해 주세요.", banner, Application.UserName)
        Loop Until argURNM <> vbNullString
    End If
    Application.UserName = argURNM
    MsgBox "이 프로그램 사용자의 이름을 '" & argURNM & "'으로 지정하였습니다.", vbInformation, banner

End Sub

'------------------------
'  등록된 사용자 체크
'------------------------
Function checkUserName(argUserNM As String, argRegisteredUser As Variant) As Boolean
    Dim userNM As Variant
       
    checkUserName = False
    For Each userNM In argRegisteredUser
        If userNM = argUserNM Then
            checkUserName = True
            Exit For
        End If
    Next userNM
    
End Function


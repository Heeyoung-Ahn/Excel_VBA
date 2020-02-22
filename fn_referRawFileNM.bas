Attribute VB_Name = "fn_referRawFileNM"
Option Explicit

'--------------------------------------------------------------------
'  원본파일의 경로포함 이름을 반환하는 함수 프로시저
'    - referRawF(원본파일이 들어있는 폴더명, 원본파일 이름)
'    - 예) referRawF("00 공통기초자료", "*교회목록*.xls*")
'--------------------------------------------------------------------
Public Function referRawFileNM(argFolderNM As String, argFileNM As String) As String
    Dim rawP As String, rawF As String
    Dim i As Integer

    For i = 1 To 24
        rawP = Chr(66 + i) & ":\" & argFolderNM & "\"
        rawF = Dir(rawP & "*" & argFileNM) '원본파일 경로포함 이름
        If Left(rawF, 1) = "~" Then
            MsgBox "파일을 다른 누군가가 열고 있습니다." & vbNewLine & _
                "확인 후 다시 진행해 주세요.", vbInformation, "파일이름 불러오기"
            Exit Function
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox "찾는 파일이 없습니다." & Space(7) & vbNewLine & _
        "확인 후 다시 진행해 주세요.", vbInformation, "파일이름 불러오기"
    Exit Function
n:
    referRawF = rawP & rawF
End Function

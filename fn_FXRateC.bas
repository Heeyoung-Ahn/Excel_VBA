Attribute VB_Name = "fn_FXRateC"
'------------------------------------------------------------------------------------------------------------------
'  환율을 계산하는 함수
'    - FXRate(화폐단위, 조회일, 조회방법: 0 원화, 1 달러화)
'------------------------------------------------------------------------------------------------------------------
Public Function FXRateC(curCode As String, Optional inqdate As Date, Optional rType As Integer = 0) As Double
    Dim m As Integer, n As Integer
    Const pbldDvCd As Integer = 0 '0 - 최종환율, 1 - 최초환율
    Application.Volatile (False)
    
    On Error Resume Next
    If inqdate = Empty Then inqdate = Date
    
    If curCode = "KRW" And (rType = Empty Or rType = 0) Then '화폐단위에 원화를 추가하고 오류 처리
        FXRateC = 1: Exit Function
    ElseIf curCode = "KRW" And rType = 1 Then
        curCode = "USD": rType = 0
    End If
    
    Select Case rType
        Case 0 '원화환산(현지화폐가 원화로는 얼마?)
            m = 9: n = 1
        Case 1 '미화환산(현지화계가 달러화로는 얼마?)
            m = 11: n = 2
        Case Else
            Exit Function
     End Select
    
    Dim WinHttp As New WinHttp.WinHttpRequest
    With WinHttp
        .Open "POST", "https://www.kebhana.com/cms/rate/wpfxd651_01i_01.do"
        .SetRequestHeader "Referer", "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_01i.do"
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send "curCd=" & UCase(curCode) & "&pbldDvCd=" & pbldDvCd & "&inqStrDt=" & Format(inqdate, "YYYYMMDD")
        .WaitForResponse
        If curCode = "JPY" Or curCode = "VND" Or curCode = "IDR" Then '엔화, 베트남 동, 인도네시아 루피는 100으로 나눠야 함
            FXRateC = (Split(Split(.ResponseText, "<td class=""txtAr"">")(m), "</td>")(0)) / 100
        Else
            FXRateC = Split(Split(.ResponseText, "<td class=""txtAr"">")(m), "</td>")(0)
        End If
    End With
    
    If FXRateC = 0 Then
        If n = 1 Then
            With WinHttp
                .Open "POST", "https://www.kebhana.com/cms/rate/wpfxd651_10i_01.do"
                .SetRequestHeader "Referer", "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_10i.do"
                .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
                .Send "inqDvCd=" & "az" & "&inqStrDt=" & Format(inqdate, "YYYYMMDD") & "&inqKindCd=" & "2"
                .WaitForResponse
                FXRateC = Split(Split(Split(.ResponseText, curCode)(1), "<td class=""txtAr"">")(n), "</td>")(0)
            End With
        Else
            With WinHttp
                .Open "POST", "https://www.kebhana.com/cms/rate/wpfxd651_10i_01.do"
                .SetRequestHeader "Referer", "https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_10i.do"
                .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
                .Send "inqDvCd=" & "az" & "&inqStrDt=" & Format(inqdate, "YYYYMMDD") & "&inqKindCd=" & "2"
                .WaitForResponse
                FXRateC = 1 / Split(Split(Split(.ResponseText, curCode)(1), "<td class=""txtAr"">")(n), "</td>")(0)
            End With
        End If
    End If
End Function


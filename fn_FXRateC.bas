Attribute VB_Name = "fn_FXRateC"
'------------------------------------------------------------------------------------------------------------------
'  ȯ���� ����ϴ� �Լ�
'    - FXRate(ȭ�����, ��ȸ��, ��ȸ���: 0 ��ȭ, 1 �޷�ȭ)
'------------------------------------------------------------------------------------------------------------------
Public Function FXRateC(curCode As String, Optional inqdate As Date, Optional rType As Integer = 0) As Double
    Dim m As Integer, n As Integer
    Const pbldDvCd As Integer = 0 '0 - ����ȯ��, 1 - ����ȯ��
    Application.Volatile (False)
    
    On Error Resume Next
    If inqdate = Empty Then inqdate = Date
    
    If curCode = "KRW" And (rType = Empty Or rType = 0) Then 'ȭ������� ��ȭ�� �߰��ϰ� ���� ó��
        FXRateC = 1: Exit Function
    ElseIf curCode = "KRW" And rType = 1 Then
        curCode = "USD": rType = 0
    End If
    
    Select Case rType
        Case 0 '��ȭȯ��(����ȭ�� ��ȭ�δ� ��?)
            m = 9: n = 1
        Case 1 '��ȭȯ��(����ȭ�谡 �޷�ȭ�δ� ��?)
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
        If curCode = "JPY" Or curCode = "VND" Or curCode = "IDR" Then '��ȭ, ��Ʈ�� ��, �ε��׽þ� ���Ǵ� 100���� ������ ��
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


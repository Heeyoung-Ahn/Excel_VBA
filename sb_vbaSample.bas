Attribute VB_Name = "sb_vbaSample"
Option Explicit

'-------------------------
'  구조체 설정
'    - 변수: 그릇
'    - 배열: 여러 그릇
'    - 구조체: 세트메뉴
'-------------------------
Type ExecuteTime
    start_time As Date
    end_time As Date
End Type

'---------------------------------------
'  시간계산 Function Procedure
'---------------------------------------
Function CalExeTime(dteStart As Date, dteEnd As Date) As String

    CalExeTime = Format(dteEnd - dteStart, "hh:nn:ss")
        
End Function

'---------------------------------------
'  배열 사용에 따른 시간 계산
'---------------------------------------
Sub checkTime1()

    Dim i As Long, k As Long
    Dim rngData() As Long
    Dim shtTask As Worksheet
    Dim calTime As ExecuteTime
    
    '//변수 설정
    ReDim rngData(50000, 100)
    Set shtTask = Sheet1
    
    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//시작시간 설정
    calTime.start_time = Time
    
    '//초기화
    shtTask.UsedRange.Delete shift:=xlUp
    
    '//배열에 계산 값 반환
    For i = 0 To 49999
        For k = 0 To 99
            rngData(i, k) = (i + 1) * (k + 1)
        Next k
    Next i
    
    '//디버깅
    Debug.Print i
    Debug.Print k
    
    '//워크시트에 배열에 저장된 값 반환
    shtTask.Range("A1").Resize(i, k).Value = rngData '엑셀의 영역범위가 배열보다 작은 건 OK, 크면 #N/A 오류값으로 채워짐
    
    '//종료시간 설정
    calTime.end_time = Time
    
    '//매크로최적화원복
     With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    '//결과보고
    MsgBox "연산에 걸린 시간: " & CalExeTime(calTime.end_time, calTime.start_time) & vbNewLine & vbNewLine & _
        "  - 시작시간: " & Format(calTime.start_time, "hh:nn:ss") & vbNewLine & _
        "  - 종료시간: " & Format(calTime.end_time, "hh:nn:ss"), vbInformation, "프로시저 실행 시간 측정"
        
End Sub

'---------------------------------------
'  워크시트 직접 입력 시 시간계산
'---------------------------------------
Sub checkTime2()

    Dim i As Long, k As Long
    Dim shtTask As Worksheet
    Dim calTime As ExecuteTime
    
    '//변수 설정
    ReDim rngData(50000, 100)
    Set shtTask = Sheet1
    
    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//시작시간 설정
    calTime.start_time = Time
    
    '//초기화
    shtTask.UsedRange.Delete shift:=xlUp
    
    '//시트에 바로 입력
    For i = 0 To 49999
        For k = 0 To 99
            shtTask.Cells(i + 1, k + 1) = (i + 1) * (k + 1)
        Next k
    Next i
       
    '//종료시간 설정
    calTime.end_time = Time
    
    '//매크로최적화원복
     With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    '//결과보고
    MsgBox "연산에 걸린 시간: " & CalExeTime(calTime.end_time, calTime.start_time) & vbNewLine & vbNewLine & _
        "  - 시작시간: " & Format(calTime.start_time, "hh:nn:ss") & vbNewLine & _
        "  - 종료시간: " & Format(calTime.end_time, "hh:nn:ss"), vbInformation, "프로시저 실행 시간 측정"
        
End Sub

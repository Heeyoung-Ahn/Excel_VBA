Attribute VB_Name = "sb_UpdateFromCommonFile"
Option Explicit
Const banner As String = "공통기초자료업데이트"
Dim MName As String
Dim tskResultCD As Integer '업데이트 결과: 0 안함, 1 완료

'--------------------
'  매크로 최적화
'--------------------
Sub Optimization()
On Error Resume Next
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
On Error GoTo 0
End Sub

'-------------------------
'  매크로 최적화 원복
'-------------------------
Sub Normal()
On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
On Error GoTo 0
End Sub

'----------------------------------
'  업데이트 체크
'    - 업데이트 진행 확인
'    - 업데이트 진행 결과 체크
'----------------------------------
Sub checkUpdate()
    MName = "업데이트할 파일이름" '설정 ★★

    If MsgBox(MName & " 자료를 공통기초자료 폴더에서 업데이트합니다. " & vbNewLine & _
        "준비되었습니까?", vbQuestion + vbYesNo, banner) = vbNo Then
        MsgBox "그럼 다시 준비하고 업데이트를 진행해 주세요.", vbInformation, banner
        Exit Sub
    End If
    
    tskResultCD = 0
    Call UpdateFromCommonFile '작업 프로시저 설정
    Call DataCleaning '찌꺼기 정리
    If tskResultCD = 1 Then
        MsgBox MName & " 자료 업데이트가 완료되었습니다." & Space(10), vbInformation, banner
    End If
End Sub

'---------------------------------------------------------------------
'  공통폴더의 공통기초자료 파일을 열어서 작업 파일 업데이트
'    - 특정폴더에 업데이트 대상 파일 유무 확인
'    - 기존파일과 업데이트하려는 파일의 구조 비교
'    - 공통기초자료로 업데이트 후 기본 서식 적용
'---------------------------------------------------------------------
Sub UpdateFromCommonFile()

On Error Resume Next
    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String, tskS As String
    Dim DB As Range
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim rawFOpen As Boolean
    Dim oldFieldNM() As String, newFieldNM() As String

    '//변수 정의
    MName = "업데이트할 파일이름" '★★
    rawS = "sheet1" '원본시트 이름 설정 ★★
    tskF = ThisWorkbook.Name '작업파일 이름 설정
    tskS = "RawData" '작업시트 이름 설정 ★★
       
    '//공통기초자료 폴더에서 업데이트 대상 파일을 찾아서 rawF에 설정
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\00 공통기초자료\" '업데이트 대상 자료의 폴더 설정 ★★
        rawF = Dir(rawP & MName) '원본파일 경로포함 이름
        If Left(rawF, 1) = "~" Then
            MsgBox MName & " 파일을 다른 누군가가 열고 있습니다.   " & vbNewLine & _
                "확인 후 다시 진행해 주세요.", vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox MName & " 파일이 업데이트하려는 폴더에 없습니다." & vbNewLine & _
        "확인 후 다시 진행해 주세요.", vbInformation, banner
    Exit Sub
n:

    '//매크로 최적화
    Call Optimization

    '//기존 작업파일 필드명 oldFieldNM 배열에 반환
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '//업데이트 대상 파일 열기
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then
            rawFOpen = True
            Exit For
        End If
    Next
    If rawFOpen = True Then
        Windows(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF, Password:="파일의 비밀번호"   ' 비밀번호 ★★
        Windows(rawF).Activate
    End If
    
    '//공통기초파일 필드명 newFieldNM 배열에 반환
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i

    '//파일 구조 점검: 필드명
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & "공통기초파일과 작업파일의 필드명이 서로 불일치합니다." & vbNewLine & _
                "확인 후 다시 진행해 주세요.", vbInformation, banner
            Windows(tskF).Activate
            GoTo m:
        End If
    Next i
    
    '//작업파일의 기초자료 초기화
    Windows(tskF).Activate
    Sheets(tskS).UsedRange.ClearContents
    
    '//공통기초자료에서 기초자료 가져오기
    Windows(rawF).Activate
    Sheets(rawS).UsedRange.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
           
    '//데이터영역설정
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '//찌꺼기 영역 삭제
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '//서식정리
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    Rows("2:2").Copy
    Rows("3:" & cntR).PasteSpecial (4)
    Application.CutCopyMode = False
    
    '//작업완료결과처리
    tskResultCD = 1
       
m:
    '//공통기초자료파일이 닫혀있었다면 다시 닫기
    If rawFOpen = False Then
        Windows(rawF).Activate
        Windows(rawF).Close SaveChanges:=False
    End If

    '//마무리
    ActiveWorkbook.Save
    
    '//매크로 최적화 원복
    Call Normal
    
End Sub

'---------------------------------
'  공통기초자료 찌꺼기 제거
'    - 0값 제거하기
'    - Trim, Clean 진행
'    - 찌거기 영역 제거
'---------------------------------
Sub DataCleaning()
    Dim tskS As String
    Dim RngData As Range, Cell As Range
    Dim cntR As Integer, cntC As Integer, i As Integer, j As Integer
    Dim data() As Variant
    
    '//매크로 최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//작업영역 설정
    tskS = "RawData" '작업시트 이름 설정 ★★
    Sheets(tskS).Activate
    Set RngData = Range("A1").CurrentRegion
    cntR = RngData.Rows.Count
    cntC = RngData.Columns.Count
    ReDim data(1 To cntR - 1, 1 To cntC)
    
    '//0값 제거, Trim, Clean
    For i = 1 To cntR - 1
        For j = 1 To cntC
            Select Case Cells(2, 1).Offset(i - 1, j - 1)
                Case 0: data(i, j) = vbNullString
                Case Else: data(i, j) = Application.WorksheetFunction.Clean(Trim(Cells(2, 1).Offset(i - 1, j - 1)))
            End Select
        Next j
    Next i
    Cells(1, 1).CurrentRegion.Offset(1).ClearContents
    Cells(2, 1).Resize(cntR - 1, cntC) = data
    
    '//찌꺼기 영역 제거
    RngData.Cells(cntR + 1, 1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp

    '//마무리
    ActiveWorkbook.Save
    
    '//매크로 최적화 원복
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
End Sub



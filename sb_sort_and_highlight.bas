Attribute VB_Name = "sb_sort_and_highlight"
Option Explicit

Sub sort_plan()
    Dim rngPlan As Range, rngA As Range
    Dim i As Integer, j As Integer
    Dim sortCase As Integer
        
    Sheet3.Activate
    '//사업우선순위
    '대상 데이터 검증
        On Error Resume Next
        i = Range("사업계획목록").Rows.Count '레코드 수
        On Error GoTo 0
        If i = 0 Then Exit Sub
    '사업우선순위 선정
        Set rngA = Sheet3.Rows(6).Find("사업우선순위", lookat:=xlWhole)
        For j = 1 To i
            Select Case rngA.Offset(j, -2) & rngA.Offset(j, -1)
                Case "상상": rngA.Offset(j).Value = "1순위"
                Case "상중": rngA.Offset(j).Value = "2순위"
                Case "중상": rngA.Offset(j).Value = "2순위"
                Case "상하": rngA.Offset(j).Value = "3순위"
                Case "중중": rngA.Offset(j).Value = "3순위"
                Case "하상": rngA.Offset(j).Value = "3순위"
                Case "중하": rngA.Offset(j).Value = "4순위"
                Case "하중": rngA.Offset(j).Value = "4순위"
                Case "하하": rngA.Offset(j).Value = "5순위"
                Case Else: rngA.Offset(j).Value = ""
            End Select
        Next j
    
    '//영역설정
    i = i + 1 '목차 포함 행 수
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A6").Resize(i, j)
    
    '//정렬 옵션 선택
    Do
        sortCase = Application.InputBox("원하는 정렬 방법을 숫자로 입력해 주세요." & vbNewLine & vbNewLine & _
                                                         "1: 사업구분 + 사업우선순위" & vbNewLine & _
                                                         "2: 사업계획코드", banner, 1, Type:=1)
    Loop While sortCase = 0 Or sortCase > 2
    
    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//정렬
    Select Case sortCase
        Case 1 '원하는 방식으로 정렬
            ActiveSheet.AutoFilterMode = False
            rngPlan.Select
            Selection.AutoFilter
            With Sheet3.AutoFilter.Sort
                .SortFields.Clear
                .SortFields.Add Key:=rngPlan.Cells(1, 15), CustomOrder:="정책사업-추진, 정책사업-일상, 행정운영경비"
                .SortFields.Add Key:=rngPlan.Cells(1, 9), Order:=xlAscending
                .SortFields.Add Key:=rngPlan.Cells(1, 11), Order:=xlAscending
                .SortFields.Add Key:=rngPlan.Cells(1, 3), Order:=xlAscending
                .Header = xlYes
                .Apply
            End With
            Call highlight_plan '조건부서식 적용
        Case 2 '코드로 정렬
            ActiveSheet.AutoFilterMode = False
            rngPlan.Select
            Selection.AutoFilter
            With Sheet3.AutoFilter.Sort
                .SortFields.Clear
                .SortFields.Add Key:=rngPlan.Cells(1, 2), Order:=xlAscending
                .Header = xlYes
                .Apply
            End With
            Call initialize_highlight_plan '조건부서식 해제
    End Select
    Cells(2, 1).Activate

    '//매크로최적화원복
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "사업계획목록 정렬이 완료되었습니다.", vbInformation, banner
End Sub

Sub highlight_plan()
    Dim rngPlan As Range, rngRow As Range, cell As Range
    Dim i As Integer, j As Integer

    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//영역설정
    i = Range("사업계획목록").Rows.Count
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A7").Resize(i, j)
    
    '//조건부서식 적용
    For Each cell In rngPlan.Resize(i, 1).Offset(0, 8)
        Select Case cell
            Case "1순위"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 6
            Case "2순위"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 36
            Case "3순위"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 19
            Case "4순위"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 15
            Case "5순위"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 48
            Case Else
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = xlNone
        End Select
    Next cell
        
    '//매크로최적화원복
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub initialize_highlight_plan()
    Dim rngPlan As Range, rngRow As Range, cell As Range
    Dim i As Integer, j As Integer

    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
'            .Calculation = xlCalculationManual
    End With
    
    '//영역설정
    i = Range("사업계획목록").Rows.Count
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A7").Resize(i, j)
    
    '//조건부 서식 지우기
    rngPlan.Interior.ColorIndex = xlNone
        
    '//매크로최적화원복
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

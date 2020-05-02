Attribute VB_Name = "sb_makeReport"
Option Explicit
Dim rngA As Range, rngB As Range, rngC As Range
Public Const banner As String = "판매현황 조회 프로그램"

'--------------------------------------
'  리포트 조회
'    - 매크로 최적화
'    - 리포트 초기화
'    - 리포트 만들기: 고급필터 활용
'    - 매크로 원복
'---------------------------------------
Sub referReport()
    '//매크로최적화
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationAutomatic
    End With
    
    Call initializeReport
    Call makeReport
    
    '//매크로최적화원복
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "리포트 조회가 완료되었습니다.", vbInformation, banner
End Sub

'----------------------
'  리포트 초기화
'    - 영역지정
'    - 입력내용 삭제
'----------------------
Sub initializeRepport()
    With Sheets("report")
        '[보호해제]
            .Unprotect Password:="12345"
        '[영역설정]
            Set rngA = .Columns("B").Find("■ 서울지역 소파 판매 목록", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("■ 광주지역 책상 판매 목록", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("■ 대전지역 침대 판매 목록", lookat:=xlWhole)
        '[입력내용초기화]
            .Range("B4").ClearContents '보고일 등 개요 초기화
        '[서울지역 소파 판매 보고서 영역 초기화]
            .Rows(rngA.Offset(2).Row & ":" & rngB.Offset(-1).Row).ClearContents
            If rngB.Row - rngA.Row > 4 Then
                .Rows(rngA.Offset(3).Row & ":" & rngB.Offset(-2).Row).Delete shift:=xlUp
            End If
        '[광주지역 책상 판매 보고서 영역 초기화]
            .Rows(rngB.Offset(2).Row & ":" & rngC.Offset(-1).Row).ClearContents
            If rngC.Row - rngB.Row > 4 Then
                .Rows(rngB.Offset(3).Row & ":" & rngC.Offset(-2).Row).Delete shift:=xlUp
            End If
        '[대전지역 침대 판매 보고서 영역 초기화]
            .Rows(rngC.Offset(2).Row & ":" & Cells(Rows.Count, 1).Row).ClearContents
            .Rows(rngC.Offset(3).Row & ":" & Cells(Rows.Count, 1).Row).Delete shift:=xlUp
        '[보호]
            .Protect Password:="12345"
        '[마무리]
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub

'-------------------------------------------------------------------
'  리포트 만들기
'    - 원본데이터 정제
'    - 고급필터 영역 설정
'      # 목록범위와 조건범위만 설정 후
'      # 실제 고급필터 시 복사위치 영역 설정
'    - 소제목 3개의 리포트 작성
'-------------------------------------------------------------------
Sub makeReport()
    Dim i As Integer
    Dim iRow As Integer, jRow As Integer
    Dim rngDB As Range, rngCriteria As Range, rngCopy As Range
    Dim cntR As Integer, cntC As Integer
    Dim rngZ As Range, cell As Range
        
    '//data adjust
    With Sheets("data")
        Set rngDB = .Range("A1").CurrentRegion
        cntR = rngDB.Rows.Count
        cntC = rngDB.Columns.Count
    End With
    Set rngZ = rngDB.Resize(1).Find(what:="판매단가", lookat:=xlWhole).Offset(1).Resize(cntR - 1, 3)
    For Each cell In rngZ
        cell.Value = Format(cell, "#,##0")
    Next cell

    '//고급필터용 영역 설정
    With Sheets("data")
        Set rngDB = .Range("A1").CurrentRegion
        Set rngCriteria = .Range("K1").CurrentRegion.Resize(1)
        Set rngCopy = .Range("N1").CurrentRegion.Resize(1)
    End With

    With Sheets("report")
        '//보호해제
            .Unprotect Password:="12345"
        '//영역설정
            Set rngA = .Columns("B").Find("■ 서울지역 소파 판매 목록", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("■ 광주지역 책상 판매 목록", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("■ 대전지역 침대 판매 목록", lookat:=xlWhole)
        '//서울지역 소파 판매 리포트 작성
            '[보고행수]
                 i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "서울", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "소파")
            '[보고서 영역 조정]
                If i > 1 Then
                    .Rows(rngA.Offset(3).Row & ":" & rngA.Offset(3).Row + i - 2).Insert shift:=xlDown
                    rngA.Offset(2).EntireRow.Copy .Range(rngA.Offset(3, -1), rngA.Offset(3 + i - 2, -1))
                End If
            '[보고서 입력]
                rngCriteria.Cells(1, 1).Offset(1).Value = "서울"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "소파"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngA.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//광주지역 책상 판매 리포트 작성
            '[보고행수]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "광주", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "책상")
            '[보고서 영역 조정]
                If i > 1 Then
                    .Rows(rngB.Offset(3).Row & ":" & rngB.Offset(3).Row + i - 2).Insert shift:=xlDown
                    rngB.Offset(2).EntireRow.Copy .Range(rngB.Offset(3, -1), rngB.Offset(3 + i - 2, -1))
                End If
            '[보고서 입력]
                rngCriteria.Cells(1, 1).Offset(1).Value = "광주"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "책상"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngB.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//대전지역 침대 판매 리포트 작성
            '[보고행수]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "대전", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "침대")
            '[보고서 영역 조정]
                If i > 1 Then
                    rngC.Offset(2).EntireRow.Copy .Range(rngC.Offset(3, -1), rngC.Offset(3 + i - 2, -1))
                End If
            '[보고서 입력]
                rngCriteria.Cells(1, 1).Offset(1).Value = "대전"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "침대"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngC.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//보고일 입력
            .Range("B4").Value = "'-보고일: " & DatePart("yyyy", Date) & "년 " & DatePart("m", Date) & "월 " & _
                DatePart("d", Date) & "일(" & Format(Date, "aaa") & ")"
        '//보호
            .Protect Password:="12345"
        '//마무리
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub



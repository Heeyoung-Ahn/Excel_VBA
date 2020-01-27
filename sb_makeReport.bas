Attribute VB_Name = "sb_makeReport"
Option Explicit
Dim rngA As Range, rngB As Range, rngC As Range

Sub initializeRepport()
    With Sheets("report")
        '[보호해제]
        .Unprotect Password:="12345"
        '[영역설정]
        Set rngA = .Columns("B").Find("■ 소제목1", lookat:=xlWhole)
        Set rngB = .Columns("B").Find("■ 소제목2", lookat:=xlWhole)
        Set rngC = .Columns("B").Find("■ 소제목3", lookat:=xlWhole)
        '[입력내용초기화]
        .Range("B4").ClearContents
        rngA.Offset(2).Resize(rngB.Row - rngA.Row - 3, 7).ClearContents
        rngB.Offset(2).Resize(rngC.Row - rngB.Row - 3, 7).ClearContents
        rngC.Offset(2).Resize(.Rows.Count - rngC.Row - 1, 7).ClearContents
        '[찌꺼기영역 제거]
        Set rngC = .Columns("B").Find("■ 소제목3", lookat:=xlWhole)
        rngC.Offset(3).Resize(Rows.Count - rngC.Row - 2, 7).Delete shift:=xlUp
        '[마무리]
        .Range("B1").Activate
        Set rngA = Nothing
        Set rngB = Nothing
        Set rngC = Nothing
        ActiveWorkbook.Save
    End With
End Sub

Sub makeReport()
    Dim i As Integer
    Dim iRow As Integer, jRow As Integer

    With Sheets("report")
        '//영역설정
            Set rngA = .Columns("B").Find("■ 소제목1", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("■ 소제목2", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("■ 소제목3", lookat:=xlWhole)
        '//소제목1 리포트 작성
            '[보고행수]
            i = 10 '소제목1 보고행수
            '[보고서 영역 조정]
            iRow = rngB.Row - rngA.Row - 3 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
                .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert shift:=xlDown
                rngA.Offset(2).Resize(1, 7).Copy .Range(rngA.Offset(3), rngA.Offset(3 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete shift:=xlUp
            End If
            '[보고서 입력]
            '고급필터 사용하여 리포트 입력
        '//소제목2 리포트 작성
            '[보고행수]
            i = 5
            '[보고서 영역 조정]
            iRow = rngC.Row - rngB.Row - 3 '현재 리포트 영역
            jRow = i - iRow '초과 리포트 영역
            If jRow > 0 Then '데이터가 제공된 영역보다 많은 경우
                .Rows(rngC.Row - 1 & ":" & rngC.Row - 1 + jRow - 1).Insert shift:=xlDown
                rngB.Offset(2).Resize(1, 7).Copy .Range(rngB.Offset(3), rngB.Offset(3 + i - 2))
            ElseIf jRow < 0 And i <> 0 Then '데이터가 제공된 영역보다 적은 경우
                .Rows(rngC.Row - 2 & ":" & rngC.Row - 1 + jRow).Delete shift:=xlUp
            ElseIf jRow < 0 And i = 0 And iRow > 1 Then '조회 데이터가 없는 경우
                .Rows(rngC.Row - 2 & ":" & rngC.Row + jRow).Delete shift:=xlUp
            End If
            '[보고서 입력]
            '고급필터 사용하여 리포트 입력
        '//소제목3 리포트 작성
            '[보고행수]
            i = 7
            '[보고서 영역 조정]
            If i > 1 Then
                rngC.Offset(2).Resize(1, 7).Copy .Range(rngC.Offset(3), rngC.Offset(3 + i - 2))
            End If
            '[보고서 입력]
            '고급필터 사용하여 리포트 입력
        '//보고일 입력
            .Range("B4").Value = "'-보고일: " & DatePart("yyyy", Date) & "년 " & DatePart("m", Date) & "월 " & _
                DatePart("d", Date) & "일(" & Format(Date, "aaa") & ")"
        '//찌꺼기 영역 제거
            i = Cells(Rows.Count, "B").End(xlUp).Row
            Cells(Rows.Count, "B").End(xlUp).Offset(1).Resize(Rows.Count - i, 7).Delete shift:=xlUp
        '//마무리
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub

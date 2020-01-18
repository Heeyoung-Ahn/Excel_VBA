Attribute VB_Name = "sb_AdvancedFilter"
Option Explicit

'---------------------------------------
'  고급필터
'    - 초기화
'    - 고급필터 영역 범위 설정
'    - 고급필터 후 정렬
'    - 고급필터 후 찌꺼기 영역 제거
'---------------------------------------
Sub AdvancedFilter()
    Dim rngDB As Range, rngCriteria As Range, rngCopy As Range
    Dim DB As Range
    Dim cntR As Integer
    
    '//기존필터 내용 초기화
    Sheet2.Range("A1").CurrentRegion.Offset(1).ClearContents
        
    '//고급필터 영역 범위 설정
    '*************************************
    '*  주의사항                                                        *
    '*   - 조건범위와 복사위치는 같은 시트에 있어야 함   *
    '*************************************
    Set rngDB = Sheets(1).Range("A1").CurrentRegion '목록범위(데이터 영역) 설정
    Set rngCriteria = Application.InputBox("조건범위를 선택한 후 확인 버튼을 누르세요!", "조건범위 선택", Type:=8) '조건범위 InputBox로 반환
    Set rngCopy = Sheets(2).Range("A1").CurrentRegion '복사위치 설정
    
    '//고급필터 실행
    rngDB.AdvancedFilter Action:=xlFilterCopy, _
        criteriarange:=rngCriteria, _
        copytorange:=rngCopy, _
        Unique:=False '고유목록만 가져오려면 True
    
    '//정렬
    Sheets(2).Activate
    ActiveSheet.AutoFilterMode = False '시트전체의 자동필터 해제
    rngCopy.Cells(1, 1).AutoFilter '복사영역에 자동필터 적용
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rngCopy.Cells(1, 1).Offset(0, 2), CustomOrder:="본부장,부장,과장,대리,사원"
        .SortFields.Add Key:=rngCopy.Cells(1, 6).Offset(0, 1), Order:=xlDescending
        .SortFields.Add Key:=rngCopy.Cells(1, 1).Offset(0, 0), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    '//찌꺼기 영역 제거
    Set DB = ActiveSheet.Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
    
    '//파일저장
    ActiveWorkbook.Save

End Sub

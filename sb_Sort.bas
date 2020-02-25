Attribute VB_Name = "sb_Sort"
Option Explicit

Sub SortSample()

    Sheet1.Activate '정렬할 시트 선택
    ActiveSheet.AutoFilterMode = False 'AutoFilterMode = False: 자동필터 해제, FilterMode = False : 필터링된 내용 해제
    Cells(1, 1).AutoFilter
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 1), CustomOrder:="본부장,부장,과장,대리,사원"
        .SortFields.Add Key:=Cells(1, 2), Order:=xlDescending
        .SortFields.Add Key:=Cells(1, 3), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    
End Sub

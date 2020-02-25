Attribute VB_Name = "sb_Sort"
Option Explicit

Sub SortSample()

    Sheet1.Activate '������ ��Ʈ ����
    ActiveSheet.AutoFilterMode = False 'AutoFilterMode = False: �ڵ����� ����, FilterMode = False : ���͸��� ���� ����
    Cells(1, 1).AutoFilter
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 1), CustomOrder:="������,����,����,�븮,���"
        .SortFields.Add Key:=Cells(1, 2), Order:=xlDescending
        .SortFields.Add Key:=Cells(1, 3), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    
End Sub

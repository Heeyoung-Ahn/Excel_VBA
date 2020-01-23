Attribute VB_Name = "sb_AdvancedFilter"
Option Explicit

'---------------------------------------
'  �������
'    - �ʱ�ȭ
'    - ������� ���� ���� ����
'    - ������� �� ����
'    - ������� �� ��� ���� ����
'---------------------------------------
Sub AdvancedFilter()
    Dim rngDB As Range, rngCriteria As Range, rngCopy As Range
    Dim DB As Range
    Dim cntR As Integer
    
    '//�������� ���� �ʱ�ȭ
    Sheet2.Range("A1").CurrentRegion.Offset(1).ClearContents
        
    '//������� ���� ���� ����
    '*************************************
    '*  ���ǻ���                                                        *
    '*   - ���ǹ����� ������ġ�� ���� ��Ʈ�� �־�� ��   *
    '*************************************
    Set rngDB = Sheets(1).Range("A1").CurrentRegion '��Ϲ���(������ ����) ����
    Set rngCriteria = Application.InputBox("���ǹ����� ������ �� Ȯ�� ��ư�� ��������!", "���ǹ��� ����", Type:=8) '���ǹ��� InputBox�� ��ȯ
    Set rngCopy = Sheets(2).Range("A1").CurrentRegion '������ġ ����
    
    '//������� ����
    rngDB.AdvancedFilter Action:=xlFilterCopy, _
        criteriarange:=rngCriteria, _
        copytorange:=rngCopy, _
        Unique:=False '������ϸ� ���������� True
    
    '//����
    Sheets(2).Activate
    ActiveSheet.AutoFilterMode = False '��Ʈ��ü�� �ڵ����� ����
    rngCopy.Cells(1, 1).AutoFilter '���翵���� �ڵ����� ����
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rngCopy.Cells(1, 1).Offset(0, 2), CustomOrder:="������,����,����,�븮,���"
        .SortFields.Add Key:=rngCopy.Cells(1, 6).Offset(0, 1), Order:=xlDescending
        .SortFields.Add Key:=rngCopy.Cells(1, 1).Offset(0, 0), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    
    '//��� ���� ����
    Set DB = ActiveSheet.Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
    
    '//��������
    ActiveWorkbook.Save

End Sub

Attribute VB_Name = "sb_sort_and_highlight"
Option Explicit

Sub sort_plan()
    Dim rngPlan As Range, rngA As Range
    Dim i As Integer, j As Integer
    Dim sortCase As Integer
        
    Sheet3.Activate
    '//����켱����
    '��� ������ ����
        On Error Resume Next
        i = Range("�����ȹ���").Rows.Count '���ڵ� ��
        On Error GoTo 0
        If i = 0 Then Exit Sub
    '����켱���� ����
        Set rngA = Sheet3.Rows(6).Find("����켱����", lookat:=xlWhole)
        For j = 1 To i
            Select Case rngA.Offset(j, -2) & rngA.Offset(j, -1)
                Case "���": rngA.Offset(j).Value = "1����"
                Case "����": rngA.Offset(j).Value = "2����"
                Case "�߻�": rngA.Offset(j).Value = "2����"
                Case "����": rngA.Offset(j).Value = "3����"
                Case "����": rngA.Offset(j).Value = "3����"
                Case "�ϻ�": rngA.Offset(j).Value = "3����"
                Case "����": rngA.Offset(j).Value = "4����"
                Case "����": rngA.Offset(j).Value = "4����"
                Case "����": rngA.Offset(j).Value = "5����"
                Case Else: rngA.Offset(j).Value = ""
            End Select
        Next j
    
    '//��������
    i = i + 1 '���� ���� �� ��
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A6").Resize(i, j)
    
    '//���� �ɼ� ����
    Do
        sortCase = Application.InputBox("���ϴ� ���� ����� ���ڷ� �Է��� �ּ���." & vbNewLine & vbNewLine & _
                                                         "1: ������� + ����켱����" & vbNewLine & _
                                                         "2: �����ȹ�ڵ�", banner, 1, Type:=1)
    Loop While sortCase = 0 Or sortCase > 2
    
    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//����
    Select Case sortCase
        Case 1 '���ϴ� ������� ����
            ActiveSheet.AutoFilterMode = False
            rngPlan.Select
            Selection.AutoFilter
            With Sheet3.AutoFilter.Sort
                .SortFields.Clear
                .SortFields.Add Key:=rngPlan.Cells(1, 15), CustomOrder:="��å���-����, ��å���-�ϻ�, ��������"
                .SortFields.Add Key:=rngPlan.Cells(1, 9), Order:=xlAscending
                .SortFields.Add Key:=rngPlan.Cells(1, 11), Order:=xlAscending
                .SortFields.Add Key:=rngPlan.Cells(1, 3), Order:=xlAscending
                .Header = xlYes
                .Apply
            End With
            Call highlight_plan '���Ǻμ��� ����
        Case 2 '�ڵ�� ����
            ActiveSheet.AutoFilterMode = False
            rngPlan.Select
            Selection.AutoFilter
            With Sheet3.AutoFilter.Sort
                .SortFields.Clear
                .SortFields.Add Key:=rngPlan.Cells(1, 2), Order:=xlAscending
                .Header = xlYes
                .Apply
            End With
            Call initialize_highlight_plan '���Ǻμ��� ����
    End Select
    Cells(2, 1).Activate

    '//��ũ������ȭ����
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "�����ȹ��� ������ �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub

Sub highlight_plan()
    Dim rngPlan As Range, rngRow As Range, cell As Range
    Dim i As Integer, j As Integer

    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//��������
    i = Range("�����ȹ���").Rows.Count
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A7").Resize(i, j)
    
    '//���Ǻμ��� ����
    For Each cell In rngPlan.Resize(i, 1).Offset(0, 8)
        Select Case cell
            Case "1����"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 6
            Case "2����"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 36
            Case "3����"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 19
            Case "4����"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 15
            Case "5����"
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = 48
            Case Else
                cell.Offset(0, -8).Resize(1, j).Interior.ColorIndex = xlNone
        End Select
    Next cell
        
    '//��ũ������ȭ����
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub initialize_highlight_plan()
    Dim rngPlan As Range, rngRow As Range, cell As Range
    Dim i As Integer, j As Integer

    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
'            .Calculation = xlCalculationManual
    End With
    
    '//��������
    i = Range("�����ȹ���").Rows.Count
    j = Range("A6").CurrentRegion.Columns.Count
    Set rngPlan = Range("A7").Resize(i, j)
    
    '//���Ǻ� ���� �����
    rngPlan.Interior.ColorIndex = xlNone
        
    '//��ũ������ȭ����
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

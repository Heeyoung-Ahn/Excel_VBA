Attribute VB_Name = "sb_makeReport"
Option Explicit
Dim rngA As Range, rngB As Range, rngC As Range
Public Const banner As String = "�Ǹ���Ȳ ��ȸ ���α׷�"

'--------------------------------------
'  ����Ʈ ��ȸ
'    - ��ũ�� ����ȭ
'    - ����Ʈ �ʱ�ȭ
'    - ����Ʈ �����: ������� Ȱ��
'    - ��ũ�� ����
'---------------------------------------
Sub referReport()
    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationAutomatic
    End With
    
    Call initializeReport
    Call makeReport
    
    '//��ũ������ȭ����
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    MsgBox "����Ʈ ��ȸ�� �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub

'-----------------------------------
'  ����Ʈ �ʱ�ȭ
'    - ��������
'    - �Է³��� ����
'    - ��� ���� ����
'------------------------------------
Sub initializeReport()
    With Sheets("report")
        '//��ȣ����
            .Unprotect Password:="12345"
        '//��������
            Set rngA = .Columns("B").Find("�� �������� ���� �Ǹ� ���", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("�� �������� å�� �Ǹ� ���", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("�� �������� ħ�� �Ǹ� ���", lookat:=xlWhole)
        '//�Է³����ʱ�ȭ
            .Range("B4").ClearContents
            rngA.Offset(2).Resize(rngB.Row - rngA.Row - 3, 7).ClearContents
            rngB.Offset(2).Resize(rngC.Row - rngB.Row - 3, 7).ClearContents
            rngC.Offset(2).Resize(.Rows.Count - rngC.Row - 1, 7).ClearContents
        '//��⿵�� ����
            rngC.Offset(3).Resize(Rows.Count - rngC.Row - 2, 7).Delete shift:=xlUp
        '//��ȣ
            .Protect Password:="12345"
        '//������
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub

'-------------------------------------------------------------------
'  ����Ʈ �����
'    - ���������� ����
'    - ������� ���� ����
'      # ��Ϲ����� ���ǹ����� ���� ��
'      # ���� ������� �� ������ġ ���� ����
'    - ������ 3���� ����Ʈ �ۼ�
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
    Set rngZ = rngDB.Resize(1).Find(what:="�ǸŴܰ�", lookat:=xlWhole).Offset(1).Resize(cntR - 1, 3)
    For Each cell In rngZ
        cell.Value = Format(cell, "#,##0")
    Next cell
    
    '//������Ϳ� ���� ����
    With Sheets("data")
        Set rngDB = .Range("A1").CurrentRegion
        Set rngCriteria = .Range("K1").CurrentRegion.Resize(1)
        Set rngCopy = .Range("N1").CurrentRegion.Resize(1)
    End With

    With Sheets("report")
        '//��ȣ����
            .Unprotect Password:="12345"
        '//��������
            Set rngA = .Columns("B").Find("�� �������� ���� �Ǹ� ���", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("�� �������� å�� �Ǹ� ���", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("�� �������� ħ�� �Ǹ� ���", lookat:=xlWhole)
        '//������1 ����Ʈ �ۼ�
            '[�������]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "����")
            '[���� ���� ����]
                iRow = rngB.Row - rngA.Row - 3 '���� ����Ʈ ����
                jRow = i - iRow '�ʰ� ����Ʈ ����
                If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
                    .Rows(rngB.Row - 1 & ":" & rngB.Row - 1 + jRow - 1).Insert shift:=xlDown
                    rngA.Offset(2).Resize(1, 7).Copy .Range(rngA.Offset(3), rngA.Offset(3 + i - 2))
                ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                    .Rows(rngB.Row - 2 & ":" & rngB.Row - 1 + jRow).Delete shift:=xlUp
                ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                    .Rows(rngB.Row - 2 & ":" & rngB.Row + jRow).Delete shift:=xlUp
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "����"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i, 7).Copy
                rngA.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//������2 ����Ʈ �ۼ�
            '[�������]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "å��")
            '[���� ���� ����]
                iRow = rngC.Row - rngB.Row - 3 '���� ����Ʈ ����
                jRow = i - iRow '�ʰ� ����Ʈ ����
                If jRow > 0 Then '�����Ͱ� ������ �������� ���� ���
                    .Rows(rngC.Row - 1 & ":" & rngC.Row - 1 + jRow - 1).Insert shift:=xlDown
                    rngB.Offset(2).Resize(1, 7).Copy .Range(rngB.Offset(3), rngB.Offset(3 + i - 2))
                ElseIf jRow < 0 And i <> 0 Then '�����Ͱ� ������ �������� ���� ���
                    .Rows(rngC.Row - 2 & ":" & rngC.Row - 1 + jRow).Delete shift:=xlUp
                ElseIf jRow < 0 And i = 0 And iRow > 1 Then '��ȸ �����Ͱ� ���� ���
                    .Rows(rngC.Row - 2 & ":" & rngC.Row + jRow).Delete shift:=xlUp
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "å��"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i, 7).Copy
                rngB.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//������3 ����Ʈ �ۼ�
            '[�������]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "ħ��")
            '[���� ���� ����]
                If i > 1 Then
                    rngC.Offset(2).Resize(1, 7).Copy .Range(rngC.Offset(3), rngC.Offset(3 + i - 2))
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "ħ��"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i, 7).Copy
                rngC.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//������ �Է�
                .Range("B4").Value = "'-������: " & DatePart("yyyy", Date) & "�� " & DatePart("m", Date) & "�� " & _
                    DatePart("d", Date) & "��(" & Format(Date, "aaa") & ")"
        '//��� ���� ����
            i = Cells(Rows.Count, "B").End(xlUp).Row
            Cells(Rows.Count, "B").End(xlUp).Offset(1).Resize(Rows.Count - i, 7).Delete shift:=xlUp
        '//��ȣ
            .Protect Password:="12345"
        '//������
            .Activate
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub


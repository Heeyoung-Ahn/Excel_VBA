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

'----------------------
'  ����Ʈ �ʱ�ȭ
'    - ��������
'    - �Է³��� ����
'----------------------
Sub initializeRepport()
    With Sheets("report")
        '[��ȣ����]
            .Unprotect Password:="12345"
        '[��������]
            Set rngA = .Columns("B").Find("�� �������� ���� �Ǹ� ���", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("�� �������� å�� �Ǹ� ���", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("�� �������� ħ�� �Ǹ� ���", lookat:=xlWhole)
        '[�Է³����ʱ�ȭ]
            .Range("B4").ClearContents '������ �� ���� �ʱ�ȭ
        '[�������� ���� �Ǹ� ���� ���� �ʱ�ȭ]
            .Rows(rngA.Offset(2).Row & ":" & rngB.Offset(-1).Row).ClearContents
            If rngB.Row - rngA.Row > 4 Then
                .Rows(rngA.Offset(3).Row & ":" & rngB.Offset(-2).Row).Delete shift:=xlUp
            End If
        '[�������� å�� �Ǹ� ���� ���� �ʱ�ȭ]
            .Rows(rngB.Offset(2).Row & ":" & rngC.Offset(-1).Row).ClearContents
            If rngC.Row - rngB.Row > 4 Then
                .Rows(rngB.Offset(3).Row & ":" & rngC.Offset(-2).Row).Delete shift:=xlUp
            End If
        '[�������� ħ�� �Ǹ� ���� ���� �ʱ�ȭ]
            .Rows(rngC.Offset(2).Row & ":" & Cells(Rows.Count, 1).Row).ClearContents
            .Rows(rngC.Offset(3).Row & ":" & Cells(Rows.Count, 1).Row).Delete shift:=xlUp
        '[��ȣ]
            .Protect Password:="12345"
        '[������]
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
        '//�������� ���� �Ǹ� ����Ʈ �ۼ�
            '[�������]
                 i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "����")
            '[���� ���� ����]
                If i > 1 Then
                    .Rows(rngA.Offset(3).Row & ":" & rngA.Offset(3).Row + i - 2).Insert shift:=xlDown
                    rngA.Offset(2).EntireRow.Copy .Range(rngA.Offset(3, -1), rngA.Offset(3 + i - 2, -1))
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "����"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngA.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//�������� å�� �Ǹ� ����Ʈ �ۼ�
            '[�������]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "å��")
            '[���� ���� ����]
                If i > 1 Then
                    .Rows(rngB.Offset(3).Row & ":" & rngB.Offset(3).Row + i - 2).Insert shift:=xlDown
                    rngB.Offset(2).EntireRow.Copy .Range(rngB.Offset(3, -1), rngB.Offset(3 + i - 2, -1))
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "å��"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngB.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//�������� ħ�� �Ǹ� ����Ʈ �ۼ�
            '[�������]
                i = Application.WorksheetFunction.CountIfs(rngDB.Cells(1.1).Offset(0, 1).Resize(cntR, 1), "����", rngDB.Cells(1.1).Offset(0, 4).Resize(cntR, 1), "ħ��")
            '[���� ���� ����]
                If i > 1 Then
                    rngC.Offset(2).EntireRow.Copy .Range(rngC.Offset(3, -1), rngC.Offset(3 + i - 2, -1))
                End If
            '[���� �Է�]
                rngCriteria.Cells(1, 1).Offset(1).Value = "����"
                rngCriteria.Cells(1, 1).Offset(1, 1).Value = "ħ��"
                rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria.CurrentRegion, copytorange:=rngCopy, Unique:=False
                rngCopy.CurrentRegion.Offset(1).Resize(i).Copy
                rngC.Offset(2).PasteSpecial Paste:=xlPasteValues
                
        '//������ �Է�
            .Range("B4").Value = "'-������: " & DatePart("yyyy", Date) & "�� " & DatePart("m", Date) & "�� " & _
                DatePart("d", Date) & "��(" & Format(Date, "aaa") & ")"
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



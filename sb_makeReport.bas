Attribute VB_Name = "sb_makeReport"
Option Explicit
Dim rngA As Range, rngB As Range, rngC As Range

Sub initializeRepport()
    With Sheets("report")
        '[��ȣ����]
        .Unprotect Password:="12345"
        '[��������]
        Set rngA = .Columns("B").Find("�� ������1", lookat:=xlWhole)
        Set rngB = .Columns("B").Find("�� ������2", lookat:=xlWhole)
        Set rngC = .Columns("B").Find("�� ������3", lookat:=xlWhole)
        '[�Է³����ʱ�ȭ]
        .Range("B4").ClearContents
        rngA.Offset(2).Resize(rngB.Row - rngA.Row - 3, 7).ClearContents
        rngB.Offset(2).Resize(rngC.Row - rngB.Row - 3, 7).ClearContents
        rngC.Offset(2).Resize(.Rows.Count - rngC.Row - 1, 7).ClearContents
        '[��⿵�� ����]
        Set rngC = .Columns("B").Find("�� ������3", lookat:=xlWhole)
        rngC.Offset(3).Resize(Rows.Count - rngC.Row - 2, 7).Delete shift:=xlUp
        '[������]
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
        '//��������
            Set rngA = .Columns("B").Find("�� ������1", lookat:=xlWhole)
            Set rngB = .Columns("B").Find("�� ������2", lookat:=xlWhole)
            Set rngC = .Columns("B").Find("�� ������3", lookat:=xlWhole)
        '//������1 ����Ʈ �ۼ�
            '[�������]
            i = 10 '������1 �������
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
            '������� ����Ͽ� ����Ʈ �Է�
        '//������2 ����Ʈ �ۼ�
            '[�������]
            i = 5
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
            '������� ����Ͽ� ����Ʈ �Է�
        '//������3 ����Ʈ �ۼ�
            '[�������]
            i = 7
            '[���� ���� ����]
            If i > 1 Then
                rngC.Offset(2).Resize(1, 7).Copy .Range(rngC.Offset(3), rngC.Offset(3 + i - 2))
            End If
            '[���� �Է�]
            '������� ����Ͽ� ����Ʈ �Է�
        '//������ �Է�
            .Range("B4").Value = "'-������: " & DatePart("yyyy", Date) & "�� " & DatePart("m", Date) & "�� " & _
                DatePart("d", Date) & "��(" & Format(Date, "aaa") & ")"
        '//��� ���� ����
            i = Cells(Rows.Count, "B").End(xlUp).Row
            Cells(Rows.Count, "B").End(xlUp).Offset(1).Resize(Rows.Count - i, 7).Delete shift:=xlUp
        '//������
            .Range("B1").Activate
            Set rngA = Nothing
            Set rngB = Nothing
            Set rngC = Nothing
            ActiveWorkbook.Save
    End With
End Sub

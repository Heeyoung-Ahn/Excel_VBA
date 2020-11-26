Attribute VB_Name = "sb_UpdateFromCommonFile"
Option Explicit
Const banner As String = "��������ڷ������Ʈ"
Dim MName As String
Dim tskS As String
Dim tskResultCD As Integer '������Ʈ ���: 0 ����, 1 �Ϸ�

'--------------------
'  ��ũ�� ����ȭ
'--------------------
Sub Optimization()
On Error Resume Next
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
On Error GoTo 0
End Sub

'-------------------------
'  ��ũ�� ����ȭ ����
'-------------------------
Sub Normal()
On Error Resume Next
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
On Error GoTo 0
End Sub

'----------------------------------
'  ������Ʈ üũ
'    - ������Ʈ ���� Ȯ��
'    - ������Ʈ ���� ��� üũ
'----------------------------------
Sub checkUpdate()
    MName = "�������Ʈ" '���� �ڡ�

    If MsgBox(MName & " �ڷḦ ��������ڷ� �������� ������Ʈ�մϴ�." & vbNewLine & _
        "�غ�Ǿ����ϱ�?", vbQuestion + vbYesNo, banner) = vbNo Then
        MsgBox "�׷� �ٽ� �غ��ϰ� ������Ʈ�� ������ �ּ���.", vbInformation, banner
        Exit Sub
    End If
    
    tskResultCD = 0
    Call UpdateFromCommonFile '�۾� ���ν��� ����
    Call DataCleaning '��� ����
    Range("A1").Activate
    If tskResultCD = 1 Then
        MsgBox MName & " �ڷ� ������Ʈ�� �Ϸ�Ǿ����ϴ�." & Space(10), vbInformation, banner
    End If
End Sub

'---------------------------------------------------------------------
'  ���������� ��������ڷ� ������ ��� �۾� ���� ������Ʈ
'    - Ư�������� ������Ʈ ��� ���� ���� Ȯ��
'    - �������ϰ� ������Ʈ�Ϸ��� ������ ���� ��
'    - ��������ڷ�� ������Ʈ �� �⺻ ���� ����
'---------------------------------------------------------------------
Sub UpdateFromCommonFile()

    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String
    Dim DB As Range
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim rawFOpen As Boolean
    Dim oldFieldNM() As String, newFieldNM() As String

    '//���� ����
    rawF = "*�������Ʈ.xls*" '�������� �̸� ���� �ڡ�
    rawS = "���" '������Ʈ �̸� ���� �ڡ�
    tskF = ThisWorkbook.Name '�۾����� �̸� ����
    tskS = "RawData" '�۾���Ʈ �̸� ���� �ڡ�
       
    '//��������ڷ� �������� ������Ʈ ��� ������ ã�Ƽ� rawF�� ����
    On Error Resume Next
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\00 ��������ڷ�\" '������Ʈ ��� �ڷ��� ���� ���� �ڡ�
        rawF = Dir(rawP & "*" & rawF) '�������� ������� �̸�
        If Left(rawF, 1) = "~" Then
            MsgBox MName & " ������ �ٸ� �������� ���� �ֽ��ϴ�.   " & vbNewLine & _
                "Ȯ�� �� �ٽ� ������ �ּ���.", vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox MName & " ������ ������Ʈ�Ϸ��� ������ �����ϴ�." & vbNewLine & _
        "Ȯ�� �� �ٽ� ������ �ּ���.", vbInformation, banner
    Exit Sub
    On Error GoTo 0
n:

    '//��ũ�� ����ȭ
    Call Optimization

    '//���� �۾����� �ʵ�� oldFieldNM �迭�� ��ȯ
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '//������Ʈ ��� ���� ����
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then
            rawFOpen = True
            Exit For
        End If
    Next
    If rawFOpen = True Then
        Workbooks(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF, Password:="12345"   ' ��й�ȣ �ڡ�
        Workbooks(rawF).Activate
    End If
    
    '//����������� �ʵ�� newFieldNM �迭�� ��ȯ
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i

    '//���� ���� ����: �ʵ��
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & "����������ϰ� �۾������� �ʵ���� ���� ����ġ�մϴ�." & vbNewLine & _
                "Ȯ�� �� �ٽ� ������ �ּ���.", vbInformation, banner
            Workbooks(tskF).Activate
            GoTo m:
        End If
    Next i
    
    '//�۾������� �����ڷ� �ʱ�ȭ
    Workbooks(tskF).Sheets(tskS).UsedRange.ClearContents
    
    '//��������ڷῡ�� �����ڷ� ��������
    Workbooks(rawF).Sheets(rawS).UsedRange.Copy
    Workbooks(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
           
    '//�����Ϳ�������
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '//��� ���� ����
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '//��������
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    Rows("2:2").Copy
    Rows("3:" & cntR).PasteSpecial (4)
    Application.CutCopyMode = False
    
    '//�۾��Ϸ���ó��
    tskResultCD = 1
       
m:
    '//��������ڷ������� �����־��ٸ� �ٽ� �ݱ�
    If rawFOpen = False Then
        Workbooks(rawF).Close SaveChanges:=False
    End If

    '//������
    ActiveWorkbook.Save
    
    '//��ũ�� ����ȭ ����
    Call Normal
    
End Sub

'---------------------------------
'  ��������ڷ� ��� ����
'    - 0�� �����ϱ�
'    - Trim, Clean ����
'    - ��ű� ���� ����
'---------------------------------
Sub DataCleaning()
    Dim RngData As Range, Cell As Range
    Dim cntR As Integer, cntC As Integer, i As Integer, j As Integer
    Dim data() As Variant
    
    '//��ũ�� ����ȭ
    Call Optimization
    
    '//�۾����� ����
    Sheets(tskS).Activate
    Set RngData = Range("A1").CurrentRegion
    cntR = RngData.Rows.Count
    cntC = RngData.Columns.Count
    ReDim data(1 To cntR - 1, 1 To cntC)
    
    '//0�� ����, Trim, Clean
    For i = 1 To cntR - 1
        For j = 1 To cntC
            Select Case Cells(2, 1).Offset(i - 1, j - 1)
                Case 0: data(i, j) = vbNullString
                Case Else: data(i, j) = Application.WorksheetFunction.Clean(Trim(Cells(2, 1).Offset(i - 1, j - 1)))
            End Select
        Next j
    Next i
    Cells(1, 1).CurrentRegion.Offset(1).ClearContents
    Cells(2, 1).Resize(cntR - 1, cntC) = data
    
    '//��� ���� ����
    RngData.Cells(cntR + 1, 1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp

    '//������
    ActiveWorkbook.Save
    
    '//��ũ�� ����ȭ ����
    Call Normal
    
End Sub

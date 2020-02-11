Attribute VB_Name = "sb_CollectData"
Option Explicit
Public Const banner As String = "���ϼ������α׷�"

'----------------------------------------------------
'  �������� ���� ������ ������ ����
'    - �����ڷ����
'    - ��������: FileDialog Property ���
'    - ���� �� ���� ���� ����
'    - ��� ���� ��ȯ�ϸ� �ڷ� ����
'    - ���� ���� �����Ͽ� �ٸ��� Pass�ϰ� �˸�
'----------------------------------------------------
Sub CollectData()

    Dim rawPath As String, rawFile As String
    Dim taskFile As String, taskSht As String
    Dim cntFile As Integer, cntC As Integer, i As Integer
    Dim oldFieldNM() As String, newFieldNM() As String
    Dim rngDB As Range
    Dim cntR As Long
    
    Application.ScreenUpdating = False
    
    '//��������
    taskFile = ThisWorkbook.Name
    taskSht = "Data" '�۾���Ʈ�̸� �ڡ�
    
    '//�������� �ʵ�� oldFieldNM �迭�� ��ȯ
    cntC = Sheets(taskSht).Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(taskSht).Range("A1").Offset(0, i).Value
    Next i
       
    '//�����ڷ� ����
    Sheets(taskSht).UsedRange.Offset(1).ClearContents
    
    '//���� ����
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        rawPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '//���� ���� ���������� �ҷ�����, ������ ������ ��ũ�� ����
    rawFile = Dir(rawPath & "*.xls*")
    If rawFile = "" Then
        MsgBox "������ ������ ������ �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
    '//���� �� ��� ���������� ��ȯ
    cntFile = 0
    Do While rawFile <> ""
        Workbooks.Open FileName:=rawPath & rawFile
        Set rngDB = ActiveSheet.Range("A1").CurrentRegion
        '���մ�� ���� �ʵ�� newFieldNM �迭�� ��ȯ
        cntC = rngDB.Columns.Count
        ReDim newFieldNM(cntC - 1)
        For i = 0 To cntC - 1
            newFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i).Value
        Next i
        '���� ��
        For i = 0 To cntC - 1
            If oldFieldNM(i) <> newFieldNM(i) Then
                MsgBox rawFile & "�� ������ ���������� ������ �ٸ��ϴ�." & vbNewLine & _
                    "�� ������ �������� �ʰ� �ǳʶݴϴ�." & vbNewLine & _
                    "���߿� Ȯ���ϼ���.", vbCritical, banner
                cntFile = cntFile - 1
                GoTo nextFile:
            End If
        Next i
        '�ڷ� ����
        rngDB.Offset(1).Resize(rngDB.Rows.Count - 1).Copy
        Workbooks(taskFile).Activate
        Sheets(taskSht).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Workbooks(rawFile).Close savechanges:=False
nextFile:
        '����
        Set rngDB = Nothing
        rawFile = Dir()
        cntFile = cntFile + 1
    Loop
    
    '//��� ����
    Set rngDB = Range("A1").CurrentRegion
    cntR = rngDB.Rows.Count
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR).Delete shift:=xlUp
    
    '//������
    Application.ScreenUpdating = True
    Range("A1").Activate
    MsgBox cntFile & "���� ���Ͽ��� �ڷ� ������ �Ϸ��Ͽ����ϴ�.", vbInformation, banner
End Sub

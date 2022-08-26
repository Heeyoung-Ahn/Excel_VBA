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

    '//��������
    Dim rawPath As String, rawFile As String, rawSht As String
    Dim taskFile As String, taskSht As String
    Dim taskFieldNM() As Variant, rawFieldNM() As Variant
    Dim cntTC As Integer, cntRC As Integer, cntR As Long, i As Integer
    Dim rngDB As Range
    Dim cntFile As Integer
    
    Application.ScreenUpdating = False
    
    '//��������
    taskFile = "���ϼ��ջ���.xlsm"
    taskSht = "Data"
    rawSht = "Sheet1"
        
    '//taskfile ���� �迭�� ��ȯ
    Set rngDB = Sheets(taskSht).Range("A1").CurrentRegion.Rows(1)
    cntTC = rngDB.Columns.Count
    ReDim taskFieldNM(1 To cntTC)
    For i = 1 To cntTC
        taskFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
    Next i
       
    '//�����ڷ� ����
    Sheets(taskSht).Range("A1").CurrentRegion.Offset(1).ClearContents
        
    '//raw folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        rawPath = .SelectedItems(1) & Application.PathSeparator
    End With
        
    '//rawfile check
    rawFile = Dir(rawPath & "*.xls*")
    If Len(rawFile) = 0 Then
        MsgBox "������ ������ ���� ������ �����ϴ�.", vbInformation, banner
        Exit Sub
    End If
    
    '//loop
    cntFile = 0
    Do
        Workbooks.Open Filename:=rawPath & rawFile
        Set rngDB = Sheets(rawSht).Range("A1").CurrentRegion.Rows(1)
        cntRC = rngDB.Columns.Count
        'rawfile ���� �迭�� ��ȯ
        ReDim rawFieldNM(1 To cntRC)
        For i = 1 To cntRC
            rawFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
        Next i
        '������1: �ʵ��
        If cntTC <> cntRC Then
            MsgBox rawFile & "�� �ʵ� ���� TaskFile�� �ʵ� ���� �ٸ��ϴ�." & vbNewLine & _
                    "���� ���Ϸ� �ǳʶݴϴ�.", vbCritical, banner
                GoTo nextFile:
        End If
        '������2: �ʵ��
        For i = 1 To cntTC
            If taskFieldNM(i) <> rawFieldNM(i) Then
                MsgBox rawFile & "�� �ʵ���� TaskFile�� �ʵ��� �ٸ��ϴ�." & vbNewLine & _
                    "���� ���Ϸ� �ǳʶݴϴ�.", vbCritical, banner
                GoTo nextFile:
            End If
        Next i
        
        '�ڷ� ����
        rngDB.CurrentRegion.Offset(1).Resize(rngDB.CurrentRegion.Rows.Count - 1).Copy
        Workbooks(taskFile).Activate
        Sheets(taskSht).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        cntFile = cntFile + 1
        
nextFile:
        '��������
        Workbooks(rawFile).Close savechanges:=False
        
        '��������
        rawFile = Dir()
    Loop Until rawFile = ""
    
    '//��� ����
    Set rngDB = Range("A1").CurrentRegion
    cntR = rngDB.Rows.Count
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
    
    Application.ScreenUpdating = True
    
    '//�۾�����
    Range("A1").Activate
    MsgBox cntFile & "���� ���� ���� �Ϸ�", vbInformation
End Sub

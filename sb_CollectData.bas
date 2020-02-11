Attribute VB_Name = "sb_CollectData"
Option Explicit
Public Const banner As String = "���ϼ������α׷�"

'----------------------------------------------
'  �������� ���� ������ ������ ����
'    - �����ڷ����
'    - ��������: FileDialog Property ���
'    - ���� �� ���� ���� ����
'    - ��� ���� ��ȯ�ϸ� �ڷ� ����
'----------------------------------------------
Sub CollectData()

    Dim rawPath As String
    Dim rawFile As String
    Dim taskFile As String
    Dim taskSht As String
    Dim cntFile As Integer
    Dim rngDB As Range
    Dim cntR As Long
    
    Application.ScreenUpdating = False
    
    '//��������
    taskFile = ThisWorkbook.Name
    taskSht = "Data" '�۾���Ʈ�̸� �ڡ�
       
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
        rngDB.Offset(1).Resize(rngDB.Rows.Count - 1).Copy
        Workbooks(taskFile).Activate
        Sheets(taskSht).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Application.DisplayAlerts = False
            Workbooks(rawFile).Close savechanges:=False
        Application.DisplayAlerts = True
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




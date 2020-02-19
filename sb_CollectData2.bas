Attribute VB_Name = "sb_CollectData2"
Option Explicit
Public Const banner As String = "���ϼ������α׷�"

'-----------------------------------------------------------------------------
'  CollectData ���ϸ�(��-"test.xlsm"), ��Ʈ��(��-"Data", [�ʵ��])
'    - �۾���Ʈ�� ���� ���� �� ó���Ǵ� �ʵ尡 �Բ� ������ ���
'    - �����ڷ� ���� �� ���մ�� �ʵ��� ���ڵ常 ����
'-----------------------------------------------------------------------------
Sub testCollectData()
    CollectData "test11.xlsm", "data", 12
End Sub

'--------------------------------------------------------------
'  �������� ���� ������ ������ ����
'    - �����ڷ����
'    - ��������: FileDialog Property ���
'    - ���� �� ���� ���� ����
'    - ��� ���� ��ȯ�ϸ� �ڷ� ����
'    - ���� ���� �����Ͽ� �ٸ��� Pass�ϰ� �˸�
'--------------------------------------------------------------
Sub CollectData(argTaskFileNM As String, argTaskShtNM As String, Optional cntTaskField As Integer = 0)

    '//��������
    Dim rawPath As String, rawFile As String
    Dim taskFieldNM() As Variant, rawFieldNM() As Variant
    Dim cntTC As Integer, cntRC As Integer, cntR As Long, i As Integer
    Dim rngDB As Range
    Dim cntFile As Integer
    
    Application.ScreenUpdating = False
            
    '//taskfile ���� �迭�� ��ȯ
    Set rngDB = Sheets(argTaskShtNM).Range("A1").CurrentRegion.Rows(1)
    If cntTaskField = 0 Then
        cntTC = rngDB.Columns.Count
    Else
        cntTC = cntTaskField
    End If
    ReDim taskFieldNM(1 To cntTC)
    For i = 1 To cntTC
        taskFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
    Next i
       
    '//�����ڷ� ����
    cntR = rngDB.Rows.Count - 1
    If cntR <> 0 Then
        Sheets(argTaskShtNM).Range("A2").Resize(cntR, cntTC).ClearContents
    End If
        
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
        Set rngDB = Sheets(1).Range("A1").CurrentRegion.Rows(1)
        cntRC = rngDB.Columns.Count
        'rawfile ���� �迭�� ��ȯ
        ReDim rawFieldNM(1 To cntRC)
        For i = 1 To cntRC
            rawFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
        Next i
        '������1: �ʵ��
        If cntTC <> cntRC Then
            MsgBox rawFile & "�� �ʵ� ���� " & argTaskFileNM & "�� �ʵ� ���� �ٸ��ϴ�." & vbNewLine & _
                    "���� ���Ϸ� �ǳʶݴϴ�.", vbCritical, banner
                GoTo nextFile:
        End If
        '������2: �ʵ��
        For i = 1 To cntTC
            If taskFieldNM(i) <> rawFieldNM(i) Then
                MsgBox rawFile & "�� �ʵ���� " & argTaskFileNM & "�� �ʵ���� �ٸ��ϴ�." & vbNewLine & _
                    "���� ���Ϸ� �ǳʶݴϴ�.", vbCritical, banner
                GoTo nextFile:
            End If
        Next i
        
        '�ڷ� ����
        rngDB.CurrentRegion.Offset(1).Resize(rngDB.CurrentRegion.Rows.Count - 1).Copy
        Workbooks(argTaskFileNM).Activate
        Sheets(argTaskShtNM).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
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
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR).Delete shift:=xlUp
    
    Application.ScreenUpdating = True
    
    '//�۾�����
    Range("A1").Activate
    MsgBox cntFile & "���� ���� ���� �Ϸ�", vbInformation
End Sub
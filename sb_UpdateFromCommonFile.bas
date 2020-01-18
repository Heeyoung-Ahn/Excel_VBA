Attribute VB_Name = "sb_UpdateFromCommonFile"
Option Explicit
Const banner As String = "��������ڷ������Ʈ"

'---------------------------------------------------------------------
'  ���������� ��������ڷ� ������ ��� �۾� ���� ������Ʈ
'
'---------------------------------------------------------------------
Sub UpdateFromCommonFile()

On Error Resume Next
    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String, tskS As String
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim DB As Range
    Dim rawFOpen As Boolean

    '//���� ����
    rawF = "���������̸�" '�ڡ�
    rawS = "������Ʈ�̸�" '�ڡ�
    tskF = ThisWorkbook.Name '�۾����� �̸�
    tskS = "�۾���Ʈ�̸�" '�ڡ�
       
    '//����DB ������ ���鼭 ������Ʈ ��� ���� ã�Ƽ� rawF�� �̸� ����
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\01 ����DB\" '���� ���� ��� �����ڡ�
        rawF = Dir(rawP & rawF) '�������� �̸� ��� ���� ����
        If Left(rawF, 1) = "~" Then '������ �ٸ� ����� ���� �ִ� ���
            MsgBox "��������ڷ� ������ �ٸ� �������� ���� ���� �ֽ��ϴ�." & vbNewLine & _
                "Ȯ�ιٶ��ϴ�." & Space(10), vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox "��������ڷ� ������ ����DB ������ �����ϴ�." & vbNewLine & _
        "Ȯ�ιٶ��ϴ�." & Space(10), vbInformation, banner
    Exit Sub
n:

    '//������Ʈ ��� ���� ����
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then rawFOpen = True
        Exit For
    Next
    If rawFOpen = True Then
        Windows(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF, Password:="������ ��й�ȣ"   '��й�ȣ�� ���� ����ڡ�
        Windows(rawF).Activate
    End If
    
    '//�۾������� �����ڷ� �ʱ�ȭ
    Windows(tskF).Activate
    Sheets(tskS).UsedRange.ClearContents
    
    '//��������ڷῡ�� �����ڷ� ��������
    Windows(rawF).Activate
    Sheets(rawS).UsedRange.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '//������Ʈ ��� ���� �ݱ�
    Windows(rawF).Close savechanges:=False '������ϰ� �ݱ�
       
    '//�����Ϳ�������
    Windows(tskF).Activate
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '//��� ���� ����
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete Shift:=xlUp
            
    '//���ʺ�����
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    
    '//2����� ��������
    Rows("2:2").Copy
    Rows("3:" & cntR).PasteSpecial (4)
    Application.CutCopyMode = False
        
    '//������
    ActiveWorkbook.Save
    
End Sub


Attribute VB_Name = "sb_DirSample"
Option Explicit

'---------------------------------
'  Dir �Լ� ����
'    - Ư�� ���� ���� ���� Ȯ��
'---------------------------------
Sub DirSample1()
    Dim strFile As String
    strFile = Dir("C:\00 ��������ڷ�\*��ȸ���*.xlsx")
    If Len(strFile) = 0 Then
        MsgBox "ã�� ������ �������� �ʽ��ϴ�.", vbCritical, "����ã��"
    Else
        MsgBox "ã�� ������ �̸��� '" & strFile & "'�Դϴ�."
    End If
End Sub

'-------------------------------
'  Dir �Լ� ����
'    - ���� �� ���� ���� ã��
'    - ���� ���� ���� ���
'    - ���� ���� �̸� ���
'-------------------------------
Sub DirSample2()
    Dim strAPath As String
    Dim strAFile As String
    Dim strFile As String
    Dim strFileSet As String
    Dim cntFile As Integer
    
    strAFile = ActiveWorkbook.FullName
    strAPath = Left(strAFile, InStrRev(strAFile, Application.PathSeparator))
    strFile = Dir(strAPath & "*.xls*")
    
    '//���� ���� Ȯ�� �� ���
    Do While strFile <> ""
        cntFile = cntFile + 1
        strFile = Dir
    Loop
    MsgBox "'" & strAPath & "' ���� �� ���� ������ ������ " & cntFile & "���Դϴ�.", vbInformation, "���ϰ�����ȸ"
    
    '//���ϸ� ���
    strFile = Dir(strAPath & "*.xls*")
    Do
        strFileSet = strFileSet & strFile & vbNewLine
        strFile = Dir
    Loop Until strFile = ""
    MsgBox strFileSet, vbInformation, "�������� �̸� ��ȸ"
End Sub

'------------------------------------------------------------------
'  Dir �Լ� ����
'    - FileDialog Property �Ӽ� �̿� ���� ����
'    - ���� ���� ��ȯ
'    - Msgbox�� ���
'------------------------------------------------------------------
Sub DirSample3()
    Dim strAPath As String
    Dim strAFile As String
    Dim strSubPath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        strAPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    strSubPath = Dir(strAPath, vbDirectory)
    Do While strSubPath <> ""
        If strSubPath <> "." And strSubPath <> ".." Then
            If (GetAttr(strAPath & strSubPath) And vbDirectory) = vbDirectory Then
                MsgBox strSubPath
            End If
        End If
        strSubPath = Dir
    Loop
End Sub


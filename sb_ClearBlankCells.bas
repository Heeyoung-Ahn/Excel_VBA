Attribute VB_Name = "sb_ClearBlankCells"
Option Explicit
Const banner As String = "Excel VBA"

'-------------------------------------------------------------
'  �������� �߻��� ��ó÷ ���̴� ���ɼ� ����
'  ���ɼ��� ������ ����� ��� ������ �����ϴ� �ڵ�
'-------------------------------------------------------------
Sub ClearBlankCells()
On Error GoTo ErrHandler:
    Dim data As Range, SelectedCell As Range, Cell As Range
    Dim cntR As Integer, cntC As Integer
    
    Set SelectedCell = Application.InputBox("�����Ͱ� �ִ� ������ �ƹ� ���̳� �����ϼ���.", banner, Type:=8)
    Set data = SelectedCell.CurrentRegion
    
    '��ó�� ���̴� ��ű� �����Ͱ� �Էµ� ���� ������ ���� �����
    For Each Cell In data
        If Len(Cell) = 0 Then
            Cell.ClearContents
        End If
    Next
    
    '��� ���� ����
    cntR = data.Rows.Count
    cntC = data.Columns.Count
    
    data.Cells(cntR + 1, 1).Resize(Rows.Count - cntR, Columns.Count).Delete
    data.Cells(1, cntC + 1).Resize(Rows.Count, Columns.Count - cntC).Delete
    
    ActiveWorkbook.Save
    Exit Sub
    
ErrHandler:
    MsgBox "������ �߻��߽��ϴ�." & Space(7) & vbNewLine & _
                  " �� ������ �߻��� ������ ĸó�Ͽ� �����ڿ��� �����ּ���." & vbNewLine & vbNewLine & _
                  "  �� �۾��� : " & Application.UserName & vbNewLine & _
                  "  �� �۾��Ͻ� : " & Now & vbNewLine & _
                  "  �� ���� �ڵ� : " & Err.Number & vbNewLine & _
                  "  �� ���� ���� : " & Err.Description & vbNewLine & _
                  "  �� ���� �ҽ� : " & Err.Source
End Sub
 

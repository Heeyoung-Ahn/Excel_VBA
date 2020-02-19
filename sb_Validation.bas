Attribute VB_Name = "sb_Validation"
Option Explicit

Sub ValidationSample()
    Dim rngDB As Range
    Dim rngA As Range
    Dim cntR As Long
        
    '//��������
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    
    '��� �������� ��������ȿ�� �˻� �����
    ActiveSheet.Cells.Validation.Delete
        
    '��ȿ���˻��� ���� ����
    Set rngA = rngDB.Rows(1).Find("�ʵ��", lookat:=xlWhole)
    With Range(rngA.Offset(1), rngA.Offset(1).Resize(cntR - 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="���1, ���2"
        .ErrorTitle = "�Է�Ȯ��"
        .ErrorMessage = "��Ͽ��� �����Ͽ� �Է��ϼ���."
    End With
        
End Sub

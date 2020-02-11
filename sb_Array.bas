Attribute VB_Name = "sb_Array"
Option Explicit

'-----------------------------------------------------------------------------
'  �迭�� �����͸� �ְ�
'  �迭 �����͸� ������ �ϳ� �ϳ� �ִ� ���
'    - �ð��� ���� �ɸ��� 1���� �迭�� ������������ �ٷ� ���� �� ����
'-----------------------------------------------------------------------------
Sub aryData1()
    
    Dim mydata(1 To 10) As Integer, intN As Integer
    
    '//�迭�� �����͸� �ְ�
    For intN = 1 To 10
        mydata(intN) = CInt(Rnd * 100)
    Next intN
    
    '//�迭 �����͸� ������ ��ȯ
    For intN = 1 To 10
        ActiveSheet.Cells(intN, 1) = mydata(intN)
    Next intN

End Sub

'--------------------------------------------------------------------------------------------------------------------------------------
'  �迭�� �����͸� �ְ�
'  �迭 �����͸� ������ �ѹ��� �ִ� ���
'    - �ð��� ���� �ɸ���, ������ �迭�� ũ�⿡ �´� ���� ������ �̸� ������ �ϰ�
'    - 1���� �迭�� �⺻������ ���ι������θ� ��ȯ�� �����Ͽ� ������������ ��ȯ�Ϸ��� transpose�Լ��� ����ؾ� ��
'--------------------------------------------------------------------------------------------------------------------------------------
Sub aryData2()

    Dim mydata(1 To 10) As Integer, intN As Integer
    Dim i As Integer, j As Integer
    
    '//�迭�� �����͸� �ְ�
    For intN = 1 To 10
        mydata(intN) = Int(Rnd * 100)
    Next intN
    
    '//�迭 �����͸� ������ ��ȯ(�������)
    ActiveSheet.Cells(1, 1).Resize(1, 10).Value = mydata
    '//�迭 �����͸� ������ ��ȯ(��������)
    ActiveSheet.Cells(1, 1).Resize(10, 1).Value = mydata 'Application.WorksheetFunction.Transpose(mydata)
    
    '//������ ������ ����ִ� ����(B2:J10)�� 99�� �Է�
    For i = 1 To 9
        For j = 1 To 9
            ActiveSheet.Cells(1, 1).Offset(i, j).Value = i * j
        Next j
    Next i
    
End Sub

'---------------------------------------------------------------------------
'  ������ ���� �����͸� �迭�� �ѹ��� ���� �ְ�
'    - �̶�, ������ ���� �������� ���������� variant�� �����ؾ� ��
'  �迭 �����͸� ������ �ѹ��� �ִ� ���
'---------------------------------------------------------------------------
Sub aryData3()

    Dim aryData() As Variant '������ ���������͸� �迭�� �ѹ��� ���� ���� ���������� Variant�� �ؾ� ��
    Dim rngDB As Range
    Dim cntR As Integer, cntC As Integer
    
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    cntC = rngDB.Columns.Count
    
    '//�����迭 ũ�� ����
    ReDim aryData(cntR - 1, cntC - 1)
    
    '//������ �ڷḦ �迭�� ��ȯ
    aryData = rngDB.Value
    
    '//�迭�� ������ ��ȯ
    ActiveSheet.Cells(20, 1).Resize(10, 10).Value = aryData
    
End Sub

'---------------------------------------------------------------------------
'  ������ ���� �����͸� �迭�� �ϳ� �� �ְ�
'    - �̶�, ������ ���� �������� ���������� ���� �������� ��밡��
'  �迭 �����͸� ������ �ѹ��� �ִ� ���
'---------------------------------------------------------------------------
Sub aryData4()

    Dim aryData() As Integer '������ �ڷḦ �ϳ� �� �迭�� ���� ���� ���������� ���� ������ ������
    Dim i As Integer, j As Integer
    Dim rngDB As Range
    Dim cntR As Integer, cntC As Integer
    Dim intR As Integer, intC As Integer
    
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    cntC = rngDB.Columns.Count
    
    '//�����迭 ũ�� ����
    ReDim aryData(cntR - 1, cntC - 1)
    
    '//���� �ڷḦ �迭��
    For i = 1 To cntR
        For j = 1 To cntC
            aryData(i - 1, j - 1) = ActiveSheet.Cells(1, 1).Offset(i - 1, j - 1).Value
        Next j
    Next i
    
    '�迭 ũ�� ������ ��ȯ
    intR = UBound(aryData, 1) - LBound(aryData, 1) + 1
    intC = UBound(aryData, 2) - LBound(aryData, 2) + 1
    
    '//�迭�� ������ ��ȯ
    ActiveSheet.Cells(20, 1).Resize(intR, intC).Value = aryData
    
End Sub

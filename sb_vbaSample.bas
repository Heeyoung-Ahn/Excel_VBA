Attribute VB_Name = "sb_vbaSample"
Option Explicit

'-------------------------
'  ����ü ����
'    - ����: �׸�
'    - �迭: ���� �׸�
'    - ����ü: ��Ʈ�޴�
'-------------------------
Type ExecuteTime
    start_time As Date
    end_time As Date
End Type

'---------------------------------------
'  �ð���� Function Procedure
'---------------------------------------
Function CalExeTime(dteStart As Date, dteEnd As Date) As String

    CalExeTime = Format(dteEnd - dteStart, "hh:nn:ss")
        
End Function

'---------------------------------------
'  �迭 ��뿡 ���� �ð� ���
'---------------------------------------
Sub checkTime1()

    Dim i As Long, k As Long
    Dim rngData() As Long
    Dim shtTask As Worksheet
    Dim calTime As ExecuteTime
    
    '//���� ����
    ReDim rngData(50000, 100)
    Set shtTask = Sheet1
    
    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//���۽ð� ����
    calTime.start_time = Time
    
    '//�ʱ�ȭ
    shtTask.UsedRange.Delete shift:=xlUp
    
    '//�迭�� ��� �� ��ȯ
    For i = 0 To 49999
        For k = 0 To 99
            rngData(i, k) = (i + 1) * (k + 1)
        Next k
    Next i
    
    '//�����
    Debug.Print i
    Debug.Print k
    
    '//��ũ��Ʈ�� �迭�� ����� �� ��ȯ
    shtTask.Range("A1").Resize(i, k).Value = rngData '������ ���������� �迭���� ���� �� OK, ũ�� #N/A ���������� ä����
    
    '//����ð� ����
    calTime.end_time = Time
    
    '//��ũ������ȭ����
     With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    '//�������
    MsgBox "���꿡 �ɸ� �ð�: " & CalExeTime(calTime.end_time, calTime.start_time) & vbNewLine & vbNewLine & _
        "  - ���۽ð�: " & Format(calTime.start_time, "hh:nn:ss") & vbNewLine & _
        "  - ����ð�: " & Format(calTime.end_time, "hh:nn:ss"), vbInformation, "���ν��� ���� �ð� ����"
        
End Sub

'---------------------------------------
'  ��ũ��Ʈ ���� �Է� �� �ð����
'---------------------------------------
Sub checkTime2()

    Dim i As Long, k As Long
    Dim shtTask As Worksheet
    Dim calTime As ExecuteTime
    
    '//���� ����
    ReDim rngData(50000, 100)
    Set shtTask = Sheet1
    
    '//��ũ������ȭ
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    '//���۽ð� ����
    calTime.start_time = Time
    
    '//�ʱ�ȭ
    shtTask.UsedRange.Delete shift:=xlUp
    
    '//��Ʈ�� �ٷ� �Է�
    For i = 0 To 49999
        For k = 0 To 99
            shtTask.Cells(i + 1, k + 1) = (i + 1) * (k + 1)
        Next k
    Next i
       
    '//����ð� ����
    calTime.end_time = Time
    
    '//��ũ������ȭ����
     With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    '//�������
    MsgBox "���꿡 �ɸ� �ð�: " & CalExeTime(calTime.end_time, calTime.start_time) & vbNewLine & vbNewLine & _
        "  - ���۽ð�: " & Format(calTime.start_time, "hh:nn:ss") & vbNewLine & _
        "  - ����ð�: " & Format(calTime.end_time, "hh:nn:ss"), vbInformation, "���ν��� ���� �ð� ����"
        
End Sub

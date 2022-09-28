Attribute VB_Name = "sb_callOfferingData"
Option Explicit

Dim MName As String
Dim tskResultCD As Integer

Sub check_update()
    MName = "������ ���� ������" '������Ʈ�� �����̸� ���� �ڡ�

    If MsgBox(MName & " �ڷḦ ��������ڷ� �������� ������Ʈ�մϴ�.           " & Chr(13) & _
        "�غ�Ǿ����ϱ�?                ", vbQuestion + vbYesNo, banner) = vbNo Then
        MsgBox "�׷� �ٽ� �غ��ϰ� ������Ʈ�� ������ �ּ���.         ", vbInformation, banner
        Exit Sub
    End If
    
    Call OfferingDBUpdate '�۾� ���ν��� ���� �ڡ�
    
    Sheets("�۾�").Activate
    
    If tskResultCD = 1 Then
        MsgBox MName & " �ڷ� ������Ʈ�� �Ϸ�Ǿ����ϴ�.    ", vbInformation, banner
    End If

End Sub

Sub OfferingDBUpdate()

    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String, tskS As String
    Dim DB As Range
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim rawFOpen As Boolean
    Dim oldFieldNM() As String, newFieldNM() As String

    '��ũ�� ����ȭ
    Call Optimization

    '���� ��ü ���� ����
    tskF = ThisWorkbook.Name '�۾����� �̸� ����
    
    '��������ڷ� �������� ������Ʈ ��� ������ ã�Ƽ� rawF�� ����
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\00 ��������ڷ�\" '������Ʈ ��� �ڷ��� ���� ���� �ڡ�
        rawF = Dir(rawP & "*20 ������ ����� ������*") '�������� �̸� ���� �ڡ�
        If Left(rawF, 1) = "~" Then
            MsgBox MName & " ������ �ٸ� �������� ���� �ֽ��ϴ�.   " & vbCrLf & _
                "Ȯ�� �� �ٽ� ������ �ּ���.            ", vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox MName & " ������ ������Ʈ�Ϸ��� ������ �����ϴ�.   " & vbCrLf & _
        "Ȯ�� �� �ٽ� ������ �ּ���.                ", vbInformation, banner
    Exit Sub
n:
    
    '���� ���� ����
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then
            rawFOpen = True
            Exit For
        End If
    Next
    If rawFOpen = True Then
        Windows(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF ', Password:="qaz1234" '��й�ȣ�� ���� ����ڡ�
        Windows(rawF).Activate
    End If
    
    '������������������������������������������
    'ù ��° ��Ʈ ���� ����
    rawS = "����ȸ ȸ�� ����(����ä����)�� ���� ��ȸ�� ���峻��" '������Ʈ �̸� ���� �ڡ�
    tskS = "t_church_offering_yyyymm_temp" '�۾���Ʈ �̸� ���� �ڡ�
    
    '���� �۾����� �ʵ�� oldFieldNM �迭�� ��ȯ
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '������ ���յ� �ʵ�� ����
    'Rows("1:1").Delete shift:=xlUp '1�࿡ ���յ� �ʵ�� ���� �ʿ�� ����ڡ�
        
    '���� ���� �ʵ�� newFieldNM �迭�� ��ȯ
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '���� ���� ����: �ʵ��
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " ������ " & tskS & " ��Ʈ�� ���� �ڷ�� �����ڷ��� �ʵ���� ���� ����ġ�մϴ�." & Chr(13) & _
                "Ȯ�� �� �ٽ� ������ �ּ���.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo m:
        End If
    Next i
            
    '�����ڷ� �ʱ�ȭ
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '���� �ڷ� ��������
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '�����Ϳ�������
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '��� ���� ����
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '��������
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0���� �ִ� ��� �����
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '�ʿ�� ���� �ڡ�

    '���ʺ�����
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

m:

    '������������������������������������������
    '�� ��° ��Ʈ ���� ����
    rawS = "����ȸ�� �����ڼ� ����" '������Ʈ �̸� ���� �ڡ�
    tskS = "t_church_offering_saint_no_yyyy" '�۾���Ʈ �̸� ���� �ڡ�
    
    '���� �۾����� �ʵ�� oldFieldNM �迭�� ��ȯ
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '������ ���յ� �ʵ�� ����
    'Rows("1:1").Delete shift:=xlUp '1�࿡ ���յ� �ʵ�� ���� �ʿ�� ����ڡ�
        
    '���� ���� �ʵ�� newFieldNM �迭�� ��ȯ
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '���� ���� ����: �ʵ��
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " ������ " & tskS & " ��Ʈ�� ���� �ڷ�� �����ڷ��� �ʵ���� ���� ����ġ�մϴ�." & Chr(13) & _
                "Ȯ�� �� �ٽ� ������ �ּ���.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo k:
        End If
    Next i
            
    '�����ڷ� �ʱ�ȭ
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '���� �ڷ� ��������
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '�����Ϳ�������
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '��� ���� ����
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '��������
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0���� �ִ� ��� �����
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '�ʿ�� ���� �ڡ�

    '���ʺ�����
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

k:

    '������������������������������������������
    '�� ��° ��Ʈ ���� ����
    rawS = "��ȸ����Ʈ" '������Ʈ �̸� ���� �ڡ�
    tskS = "t_church_disp_key_info_temp" '�۾���Ʈ �̸� ���� �ڡ�
    
    '���� �۾����� �ʵ�� oldFieldNM �迭�� ��ȯ
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '������ ���յ� �ʵ�� ����
    'Rows("1:1").Delete shift:=xlUp '1�࿡ ���յ� �ʵ�� ���� �ʿ�� ����ڡ�
        
    '���� ���� �ʵ�� newFieldNM �迭�� ��ȯ
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '���� ���� ����: �ʵ��
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " ������ " & tskS & " ��Ʈ�� ���� �ڷ�� �����ڷ��� �ʵ���� ���� ����ġ�մϴ�." & Chr(13) & _
                "Ȯ�� �� �ٽ� ������ �ּ���.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo s:
        End If
    Next i
            
    '�����ڷ� �ʱ�ȭ
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '���� �ڷ� ��������
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '�����Ϳ�������
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '��� ���� ����
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '��������
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0���� �ִ� ��� �����
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '�ʿ�� ���� �ڡ�

    '���ʺ�����
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

s:

    '���� ������ �����־��ٸ� �ٽ� �ݱ�
    If rawFOpen = False Then
        Windows(rawF).Activate
        Windows(rawF).Close SaveChanges:=False
    End If

    '������
    ActiveWorkbook.Save

    '��ũ�� ����ȭ ����
    Call Normal
    
End Sub






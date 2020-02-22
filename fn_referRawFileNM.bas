Attribute VB_Name = "fn_referRawFileNM"
Option Explicit

'--------------------------------------------------------------------
'  ���������� ������� �̸��� ��ȯ�ϴ� �Լ� ���ν���
'    - referRawF(���������� ����ִ� ������, �������� �̸�)
'    - ��) referRawF("00 ��������ڷ�", "*��ȸ���*.xls*")
'--------------------------------------------------------------------
Public Function referRawFileNM(argFolderNM As String, argFileNM As String) As String
    Dim rawP As String, rawF As String
    Dim i As Integer

    For i = 1 To 24
        rawP = Chr(66 + i) & ":\" & argFolderNM & "\"
        rawF = Dir(rawP & "*" & argFileNM) '�������� ������� �̸�
        If Left(rawF, 1) = "~" Then
            MsgBox "������ �ٸ� �������� ���� �ֽ��ϴ�." & vbNewLine & _
                "Ȯ�� �� �ٽ� ������ �ּ���.", vbInformation, "�����̸� �ҷ�����"
            Exit Function
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox "ã�� ������ �����ϴ�." & Space(7) & vbNewLine & _
        "Ȯ�� �� �ٽ� ������ �ּ���.", vbInformation, "�����̸� �ҷ�����"
    Exit Function
n:
    referRawF = rawP & rawF
End Function

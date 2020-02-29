Attribute VB_Name = "fn_checkDoubleInput"
Option Explicit

'--------------------------------------------------------------------------------------------------
'  Ư�� �ʵ� �ߺ� ����
'    - checkDoubleInput(�ʵ��, ������, DB���̺��, ��������, ����������) True: �ߺ�
'    - ������ ��� ���� �����͸� �־ �ߺ�üũ ��ȸ
'--------------------------------------------------------------------------------------------------
Public Function checkDoubleInput(fieldNM As String, data As Variant, _
                                                    tableNM As String, formNM As String, Optional ByVal beforeData As Variant = Empty) As Boolean
    Dim strSQL As String
    Dim cntRecord As Integer
    
    '//Ư�� �ʵ忡 Ư�� ������ ���� ��ȯ
    Call connectTaskDB
    strSQL = "SELECT COUNT(" & fieldNM & ") record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM & " = " & SText(data) & ";"
    callDBtoRS "checkDoubleInput", tableNM, strSQL, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    Call disconnectALL
    
    '//�ߺ� �Է� ����
    If beforeData <> Empty And beforeData = data Then Exit Function '//������ ��� ���� �����Ϳ� �����ؼ� ���
    If cntRecord >= 1 Then
        checkDoubleInput = True
    Else
        checkDoubleInput = False
    End If
End Function

'----------------------------------------------------------------------------------------------------------------------------
'  ���� ������ �ߺ� ����
'    - checkDoubleInput(����������, �ʵ��1, �ʵ��2, ������1, ������2, DB���̺��, �������̸�) True: �ߺ�
'----------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput2(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSQL As String
    Dim cntRecord As Integer
    
    '//Ư�� �ʵ忡 Ư�� ������ ���� ��ȯ
    Call connectTaskDB
    strSQL = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM1 & " = " & SText(Data1) & " AND " & _
                  fieldNM2 & " = " & SText(Data2) & ";"
    callDBtoRS "checkDoubleInput2", tableNM, strSQL, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    Call disconnectALL
    
    '//�ߺ� �Է� ����
    Select Case dataType
        Case 1 '//�ű��Է�
            If cntRecord > 0 Then
                checkDoubleInput2 = True
            Else
                checkDoubleInput2 = False
            End If
        Case 2 '//�����Է�
            If cntRecord >= 2 Then
                checkDoubleInput2 = True
            Else
                checkDoubleInput2 = False
            End If
        Case 4 '//��������
            checkDoubleInput2 = False
    End Select
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------
'  �Ⱓ ���� ������ �ߺ� ����
'    - checkDoubleInput3( ����������, �ʵ��1, �ʵ��2, ������1, ������2, ������, ������, DB���̺��, �������̸�) True: �ߺ�
'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput3(dataType As Integer, fieldNM1 As String, fieldNM2 As String, Data1 As Variant, Data2 As Variant, _
                                                      start_dt As Date, end_dt As Date, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSQL As String
    Dim cntRecord As Integer
    
    '//Ư�� �ʵ忡 Ư�� ������ ���� ��ȯ
    Call connectTaskDB
    strSQL = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM1 & " = " & SText(Data1) & " AND " & _
                  fieldNM2 & " = " & SText(Data2) & " AND " & _
                  "start_dt <= " & SText(end_dt) & " AND " & _
                  "end_dt >= " & SText(start_dt) & ";"
    callDBtoRS "checkDoubleInput3", tableNM, strSQL, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    
    Call disconnectALL
    
    '//�ߺ� �Է� ����
    Select Case dataType
        Case 1 '//�ű��Է�
            If cntRecord > 0 Then
                checkDoubleInput3 = True
            Else
                checkDoubleInput3 = False
            End If
        Case 2 '//�����Է�
            If cntRecord > 1 Then
                checkDoubleInput3 = True
            Else
                checkDoubleInput3 = False
            End If
        Case 4 '//��������
            checkDoubleInput3 = False
    End Select
End Function

'-------------------------------------------------------------------------------------------------------------------------
'  �Ⱓ ������ �ߺ� ����
'    - checkDoubleInput4( ����������, �ʵ��, ������, ������, ������, DB���̺��, �������̸�) True: �ߺ�
'-------------------------------------------------------------------------------------------------------------------------
Public Function checkDoubleInput4(dataType As Integer, fieldNM As String, data As Variant, _
                                                      start_dt As Date, end_dt As Date, _
                                                      tableNM As String, formNM As String) As Boolean
    Dim strSQL As String
    Dim cntRecord As Integer
    
    '//Ư�� �ʵ忡 Ư�� ������ ���� ��ȯ
    Call connectTaskDB
    strSQL = "SELECT COUNT(*) record_cnt " & _
                  "FROM " & tableNM & " " & _
                  "WHERE " & fieldNM & " = " & SText(data) & " AND " & _
                  "start_dt <= " & SText(end_dt) & " AND " & _
                  "end_dt >= " & SText(start_dt) & ";"
    callDBtoRS "checkDoubleInput4", tableNM, strSQL, formNM
    If rs.EOF = True Then
        cntRecord = 0
    Else
        cntRecord = rs("record_cnt").Value
    End If
    
    Call disconnectALL
    
    '//�ߺ� �Է� ����
    Select Case dataType
        Case 1 '//�ű��Է�
            If cntRecord > 0 Then
                checkDoubleInput4 = True
            Else
                checkDoubleInput4 = False
            End If
        Case 2 '//�����Է�
            If cntRecord > 1 Then
                checkDoubleInput4 = True
            Else
                checkDoubleInput4 = False
            End If
        Case 4 '//��������
            checkDoubleInput4 = False
    End Select
End Function


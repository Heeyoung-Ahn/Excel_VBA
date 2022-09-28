Attribute VB_Name = "sb_inserOfferingtDataToDB"
Option Explicit

'---------------------------------------------------
'  ���� �ڷḦ DB�� Insert
'    - Table �ʱ�ȭ
'    - ���� ���ڵ� �ϳ��� Insert
'----------------------------------------------------

Sub before_push_excel()
'���� ������ ���ε� �����۾�

    Dim strSQL As String
    Dim affectedCount As Integer, affectedCountN As Integer
    Dim Fields() As Variant
    Dim Values() As Variant
    Dim cntField As Integer, cntRecord As Integer, i As Integer, j As Integer
    
    connectTaskDB
    
    '//���ڵ庰�� Insert
    strSQL = "CALL before_push_excel;"
        affectedCountN = excuteSQL(strSQL)
        affectedCount = affectedCountN
    disconnectALL
'    MsgBox "���� ������ DB ���ε� �����۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, banner

End Sub

Sub after_push_excel()
'���� ������ ���ε� �����۾�

    Dim strSQL As String
    Dim affectedCount As Long, affectedCountN As Long
    Dim Fields() As Variant
    Dim Values() As Variant
    Dim cntField As Integer, cntRecord As Integer, i As Integer, j As Integer
    
    connectTaskDB
    
    '//���ڵ庰�� Insert
    strSQL = "CALL after_push_excel ;"
        affectedCountN = excuteSQL(strSQL)
        affectedCount = affectedCountN
    disconnectALL
    MsgBox "���� ������ DB ���ε� �����۾��� �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub


Sub t_church_offering_insert()
'���� ����ݾ� DB �ݿ�

    Dim strSQL As String
    Dim affectedCount As Integer, affectedCountN As Integer
    Dim Fields() As Variant
    Dim Values() As Variant
    Dim cntField As Integer, cntRecord As Integer, i As Integer, j As Integer
    
    connectTaskDB
      
    '/�迭 ũ�� ����
    Sheets("t_church_offering_yyyymm_temp").Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Fields(cntField - 1)
    ReDim Values(cntRecord - 1, cntField - 1)
    
    '//�ʵ�� Fields �迭�� ����
    For i = 0 To cntField - 1
        Fields(i) = Range("A1").Offset(0, i)
    Next i
    
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert
    For i = 0 To cntRecord - 1
        strSQL = "INSERT INTO regular_income.t_church_offering_yyyymm_temp(CHURCH_KEY_NO, CHURCH_NM_KOR, CHURCH_GB_NM_KOR, MANAGE_CHURCH_NM_KOR, OFFERING_GB_NM_KOR, YYYYMM_DT, OFFERING_AMT, WMC_CHARGE_NM) VALUES(" & _
                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
                       IIf((IsNull(Values(i, 7)) Or Values(i, 7) = ""), "''", (IIf(IsNumeric(Values(i, 7)), Values(i, 7), SText(Values(i, 7))))) & ");"
        affectedCountN = excuteSQL(strSQL)
        affectedCount = affectedCount + affectedCountN
    Next i
    disconnectALL
'    MsgBox affectedCount & "���� ���嵥���� ���ε尡 �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub

Sub t_church_offering_saint_no_insert()
'���� ���强���� DB �ݿ�
'��ȸ���� DB �ݿ�

    Dim strSQL As String
    Dim affectedCount As Integer, affectedCountN As Integer
    Dim Fields() As Variant
    Dim Values() As Variant
    Dim cntField As Integer, cntRecord As Integer, i As Integer, j As Integer
    
    connectTaskDB
    
    '/�迭 ũ�� ����
    Sheets("t_church_offering_saint_no_yyyy").Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Fields(cntField - 1)
    ReDim Values(cntRecord - 1, cntField - 1)
    
    '//�ʵ�� Fields �迭�� ����
    For i = 0 To cntField - 1
        Fields(i) = Range("A1").Offset(0, i)
    Next i
    
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert
    For i = 0 To cntRecord - 1
        strSQL = "INSERT INTO regular_income.t_church_offering_saint_no_yyyymm_temp(YYYYMM_DT, CHURCH_KEY_NO, CHURCH_NM_KOR, CHURCH_GB_NM_KOR, MANAGE_CHURCH_NM_KOR, SAINT_NO, WMC_CHARGE_NM) VALUES(" & _
                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & ");"
        affectedCountN = excuteSQL(strSQL)
        affectedCount = affectedCount + affectedCountN
    Next i
    disconnectALL
'    MsgBox affectedCount & "���� ���强���� ���ε尡 �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub

Sub t_church_disp_key_info_insert()
'��ȸ���� DB �ݿ�

    Dim strSQL As String
    Dim affectedCount As Integer, affectedCountN As Integer
    Dim Fields() As Variant
    Dim Values() As Variant
    Dim cntField As Integer, cntRecord As Integer, i As Integer, j As Integer
    
    connectTaskDB
    
    '/�迭 ũ�� ����
    Sheets("t_church_disp_key_info_temp").Activate
    cntField = Range("A1").CurrentRegion.Columns.Count
    cntRecord = Range("A1").CurrentRegion.Rows.Count - 1
    ReDim Fields(cntField - 1)
    ReDim Values(cntRecord - 1, cntField - 1)
    
    '//�ʵ�� Fields �迭�� ����
    For i = 0 To cntField - 1
        Fields(i) = Range("A1").Offset(0, i)
    Next i
    
    '//DB�� �̵��� �ڷ� Values �迭�� ����
    For i = 0 To cntField - 1
        For j = 0 To cntRecord - 1
            Values(j, i) = Range("A2").Offset(j, i)
        Next j
    Next i
    
    '//���ڵ庰�� Insert
    For i = 0 To cntRecord - 1
        strSQL = "INSERT INTO regular_income.t_church_disp_key_info_temp(CHURCH_GB_NM_KOR,CHURCH_KEY_NO,SORT_ORDER,OPERATION_GB,STARTING_DATE, " & _
                       "TIME_STAMP,CLOSING_DATE,CHURCH_NM_KOR,MANAGE_CHURCH_NM_KOR,COUNTRY_NM_KOR,DISP_CHURCH_GB," & _
                       "DISP_CHURCH_GB_NM_KOR,DISP_CHURCH_MANAGER_NM_KR,DUTY_NM,POSITION_NM,LAST_EMPLOYEE_NM,COMMENTS) VALUES(" & _
                       IIf((IsNull(Values(i, 0)) Or Values(i, 0) = ""), "''", (IIf(IsNumeric(Values(i, 0)), Values(i, 0), SText(Values(i, 0))))) & "," & _
                       IIf((IsNull(Values(i, 1)) Or Values(i, 1) = ""), "''", (IIf(IsNumeric(Values(i, 1)), Values(i, 1), SText(Values(i, 1))))) & "," & _
                       IIf((IsNull(Values(i, 2)) Or Values(i, 2) = ""), "''", (IIf(IsNumeric(Values(i, 2)), Values(i, 2), SText(Values(i, 2))))) & "," & _
                       IIf((IsNull(Values(i, 3)) Or Values(i, 3) = ""), "''", (IIf(IsNumeric(Values(i, 3)), Values(i, 3), SText(Values(i, 3))))) & "," & _
                       IIf((IsNull(Values(i, 4)) Or Values(i, 4) = ""), "''", (IIf(IsNumeric(Values(i, 4)), Values(i, 4), SText(Values(i, 4))))) & "," & _
                       IIf((IsNull(Values(i, 5)) Or Values(i, 5) = ""), "''", (IIf(IsNumeric(Values(i, 5)), Values(i, 5), SText(Values(i, 5))))) & "," & _
                       IIf((IsNull(Values(i, 6)) Or Values(i, 6) = ""), "''", (IIf(IsNumeric(Values(i, 6)), Values(i, 6), SText(Values(i, 6))))) & "," & _
                       IIf((IsNull(Values(i, 7)) Or Values(i, 7) = ""), "''", (IIf(IsNumeric(Values(i, 7)), Values(i, 7), SText(Values(i, 7))))) & "," & _
                       IIf((IsNull(Values(i, 8)) Or Values(i, 8) = ""), "''", (IIf(IsNumeric(Values(i, 8)), Values(i, 8), SText(Values(i, 8))))) & "," & _
                       IIf((IsNull(Values(i, 9)) Or Values(i, 9) = ""), "''", (IIf(IsNumeric(Values(i, 9)), Values(i, 9), SText(Values(i, 9))))) & "," & _
                       IIf((IsNull(Values(i, 10)) Or Values(i, 10) = ""), "''", (IIf(IsNumeric(Values(i, 10)), Values(i, 10), SText(Values(i, 10))))) & "," & _
                       IIf((IsNull(Values(i, 11)) Or Values(i, 11) = ""), "''", (IIf(IsNumeric(Values(i, 11)), Values(i, 11), SText(Values(i, 11))))) & "," & _
                       IIf((IsNull(Values(i, 12)) Or Values(i, 12) = ""), "''", (IIf(IsNumeric(Values(i, 12)), Values(i, 12), SText(Values(i, 12))))) & "," & _
                       IIf((IsNull(Values(i, 13)) Or Values(i, 13) = ""), "''", (IIf(IsNumeric(Values(i, 13)), Values(i, 13), SText(Values(i, 13))))) & "," & _
                       IIf((IsNull(Values(i, 14)) Or Values(i, 14) = ""), "''", (IIf(IsNumeric(Values(i, 14)), Values(i, 14), SText(Values(i, 14))))) & "," & _
                       IIf((IsNull(Values(i, 15)) Or Values(i, 15) = ""), "''", (IIf(IsNumeric(Values(i, 15)), Values(i, 15), SText(Values(i, 15))))) & "," & _
                       IIf((IsNull(Values(i, 16)) Or Values(i, 16) = ""), "''", (IIf(IsNumeric(Values(i, 16)), Values(i, 16), SText(Values(i, 16))))) & ");"
        affectedCountN = excuteSQL(strSQL)
        affectedCount = affectedCount + affectedCountN
    Next i
    disconnectALL
'    MsgBox affectedCount & "���� ��ȸ���� ���ε尡 �Ϸ�Ǿ����ϴ�.", vbInformation, banner
End Sub








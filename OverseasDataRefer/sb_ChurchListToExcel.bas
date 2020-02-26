Attribute VB_Name = "sb_ChurchListToExcel"
Option Explicit

'--------------------------
'  복음구획DB 엑셀 반환
'--------------------------
Sub ChurchListtoExcel()
    
    Dim tableNM As String, dbNM As String
    Dim strSQL As String
    Dim i As Integer
    Dim fileNM As String
    Dim fileSNM As String
    
    '//로그인체크
    If checkLogin = 0 Then
        MsgBox "먼저 로그인 해주세요." & Space(10), vbInformation, banner
        Exit Sub
    End If
    
    '//db명 설정
    tableNM = "overseas.v_churches" '//db명.테이블명 - 수정★★
    dbNM = "overseas" '//수정★★
    
    '//DB연결
    connectTaskDB
    
    '//Select문-수정★★
    strSQL = "SELECT * FROM " & tableNM & " WHERE `담당부서` = " & SText(user_dept) & ";"
    
    '//SQL문 실행하고 조회된 자료를 레코드셋에 담음
    callDBtoRS "gospelDBtoExcel", tableNM, strSQL, , "교회리스트엑셀반환"
    If rs.EOF = True Then
        MsgBox "조회 조건에 맞는 자료가 없습니다.", vbInformation, banner
        disconnectALL
        Exit Sub
    End If
        
    '//엑셀로 자료 내보내기
    Optimization
    Sheet3.Cells(1, 1).CurrentRegion.Offset(1).Delete shift:=xlUp
    Sheet3.Activate
    For i = 0 To rs.Fields.Count - 1
        Cells(1, 1).Offset(0, i).Value = rs.Fields(i).Name
    Next i
    Cells(2, 1).CopyFromRecordset rs
    Cells(1.1).CurrentRegion.Columns.AutoFit
    '정렬
    ActiveSheet.AutoFilterMode = False
    Cells(1, 1).AutoFilter
    With ActiveSheet.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Cells(1, 13), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    ActiveSheet.AutoFilterMode = False
    ActiveWorkbook.Save
    Normal
    
    Cells(2, 1).Select
    
    '//결과보고, 마무리
    '로그기록
    strSQL = "INSERT INTO common.logs(procedure_nm, table_nm, sql_script, error_cd, job_nm, affectedCount, user_id) " & _
                  "Values('ChurchListtoExcel', " & SText(tableNM) & ", " & SText(strSQL) & ", 0, '교회리스트엑셀반환', " & rs.RecordCount & ", " & user_id & ");"
    executeSQL "writeLog", "common.logs", strSQL, , "로그기록"
    disconnectALL
    '결과보고
    MsgBox "교회리스트 조회가 완료되었습니다.", vbInformation, banner
End Sub




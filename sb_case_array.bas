Attribute VB_Name = "sb_case_array"
Option Explicit

Sub aryData1()
    
    Dim mydata(1 To 10) As Integer, intN As Integer
    
    '//배열에 데이터를 넣고
    For intN = 1 To 10
        mydata(intN) = CInt(Rnd * 100)
    Next intN
    
    '//배열 데이터를 엑셀에 반환
    For intN = 1 To 10
        ActiveSheet.Cells(intN, 1) = mydata(intN)
    Next intN

End Sub

Sub aryData2()

    Dim mydata(1 To 10) As Integer, intN As Integer
    Dim i As Integer, j As Integer
    
    '//배열에 데이터를 넣고
    For intN = 1 To 10
        mydata(intN) = Int(Rnd * 100)
    Next intN
    
    '//배열 데이터를 엑셀에 반환(수평방향)
    ActiveSheet.Cells(1, 1).Resize(1, 10).Value = mydata
    '//배열 데이터를 엑셀에 반환(수직방향)
    ActiveSheet.Cells(1, 1).Resize(10, 1).Value = Application.WorksheetFunction.Transpose(mydata)
    
    '//데이터 범위의 비어있는 영역(B2:J10)에 99단 입력
    For i = 1 To 9
        For j = 1 To 9
            ActiveSheet.Cells(1, 1).Offset(i, j).Value = i * j
        Next j
    Next i
    
End Sub

Sub aryData3()

    Dim aryData() As Variant '엑셀의 영역데이터를 배열에 한번에 넣을 때는 데이터형을 Variant로 해야 함
    Dim rngDB As Range
    Dim cntR As Integer, cntC As Integer
    
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    cntC = rngDB.Columns.Count
    
    '//동적배열 크기 지정
    ReDim aryData(cntR - 1, cntC - 1)
    
    '//엑셀의 자료를 배열로 반환
    aryData = rngDB.Value
    
    '//배열을 엑셀에 반환
    ActiveSheet.Cells(20, 1).Resize(10, 10).Value = aryData
    
End Sub

Sub aryData4()

    Dim aryData() As Integer '엑셀의 자료를 하나 씩 배열에 넣을 때는 데이터형을 실제 데이터 형으로
    Dim i As Integer, j As Integer
    Dim rngDB As Range
    Dim cntR As Integer, cntC As Integer
    Dim intR As Integer, intC As Integer
    
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    cntC = rngDB.Columns.Count
    
    '//동적배열 크기 지정
    ReDim aryData(cntR - 1, cntC - 1)
    
    '//엑셀의 자료를 배열로
    For i = 1 To cntR
        For j = 1 To cntC
            aryData(i - 1, j - 1) = ActiveSheet.Cells(1, 1).Offset(i - 1, j - 1).Value
        Next j
    Next i
    
    '배열 크기 변수에 반환
    intR = UBound(aryData, 1) - LBound(aryData, 1) + 1
    intC = UBound(aryData, 2) - LBound(aryData, 2) + 1
    
    '//배열을 엑셀에 반환
    ActiveSheet.Cells(20, 1).Resize(intR, intC).Value = aryData
    
End Sub

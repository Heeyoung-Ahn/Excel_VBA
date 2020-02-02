Attribute VB_Name = "sb_LoopSample"
Option Explicit

Sub ForNextDemo1()
    Dim i As Long
    Dim lngEven As Long
    
    lngEven = 0
    Sheet1.Range("A1").Value = 10
    For i = 1 To 9
        Sheet1.Range("A1").Offset(i).Value = 10 - i
    Next i
End Sub

Sub ForNextDemo2()
    Dim i As Long
    Dim lngEven As Long
    
    lngEven = 0
    For i = 2 To 20 Step 2
        lngEven = lngEven + i
    Next i
    MsgBox "2부터 20까지의 숫자 중 짝수의 합계는 " & lngEven & "입니다."
End Sub

Sub ForEachNextDemo3()
    Dim sht As Worksheet
    
    For Each sht In Worksheets
        If UCase(sht.Name) = UCase("sheet1") Then
            MsgBox "해당 시트가 존재합니다."
            Exit Sub
        End If
    Next sht
     MsgBox "해당 시트가 존재하지 않습니다."
End Sub

Sub ForEachNextDemo1()
    Dim lngSum As Long
    Dim rng As Range
    Dim rngDB As Range
    
    Set rngDB = Range("A1").CurrentRegion
    lngSum = 0
    For Each rng In rngDB
        If rng.Value > 0 Then
            lngSum = lngSum + rng.Value
        End If
    Next rng
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Value = lngSum
End Sub

Sub DoUntilDemo()
    Dim rngDB As Range, rngA As Range
    Dim cntR As Integer, cntC As Integer, i As Integer
            
    '//영역설정
    Set rngDB = Sheets("DB").UsedRange
    cntR = rngDB.Rows.Count
    cntC = rngDB.Columns.Count
    
    '//사원명을 이름과 직책으로 분리
    Set rngA = rngDB.Resize(1).Find("사원명", lookat:=xlWhole)
    '[직책필드추가]
    rngA.Offset(0, 1).EntireColumn.Insert
    rngA.Offset(0, 1).Value = "직책"
    '[사원명, 직책 데이터 분리]
    i = 1
    Do
        rngA.Offset(i, 1).Value = Right(rngA.Offset(i).Value, Len(rngA.Offset(i)) - InStr(rngA.Offset(i).Value, " "))
        rngA.Offset(i).Value = Left(rngA.Offset(i).Value, InStr(rngA.Offset(i).Value, " ") - 1)
        i = i + 1
    Loop Until rngA.Offset(i) = vbNullString
End Sub


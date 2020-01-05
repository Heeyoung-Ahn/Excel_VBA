Attribute VB_Name = "sb_DataCleaning"
Option Explicit

'-----------------------------------
'  어드민 데이터 찌꺼기 제거
'    - 0값 제거하기
'    - Trim, Clean 진행
'-----------------------------------
Sub DataCleaning()

    Dim RngData As Range, Cell As Range
    Dim cntR As Integer, cntC As Integer, i As Integer, j As Integer
    Dim data() As Variant
    
    Call Optimization
    
    With Sheets("RawData")
        .Activate
        Set RngData = .[a1].CurrentRegion
        cntR = RngData.Rows.Count
        cntC = RngData.Columns.Count
        ReDim data(1 To cntR - 1, 1 To cntC)
        
        '0값 제거, Trim, Clean
        For i = 1 To cntR - 1
            For j = 1 To cntC
                Select Case Cells(2, 1).Offset(i - 1, j - 1)
                    Case 0: data(i, j) = ""
                    Case Else: data(i, j) = Application.WorksheetFunction.Clean(Trim(Cells(2, 1).Offset(i - 1, j - 1)))
                End Select
            Next j
        Next i
        Cells(1, 1).CurrentRegion.Offset(1).ClearContents
        Cells(2, 1).Resize(cntR - 1, cntC) = data
    End With

    Call Normal
    
End Sub

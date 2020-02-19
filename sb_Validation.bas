Attribute VB_Name = "sb_Validation"
Option Explicit

Sub ValidationSample()
    Dim rngDB As Range
    Dim rngA As Range
    Dim cntR As Long
        
    '//영역설정
    Set rngDB = ActiveSheet.Cells(1, 1).CurrentRegion
    cntR = rngDB.Rows.Count
    
    '모든 영역에서 데이터유효성 검사 지우기
    ActiveSheet.Cells.Validation.Delete
        
    '유효성검사할 영역 설정
    Set rngA = rngDB.Rows(1).Find("필드명", lookat:=xlWhole)
    With Range(rngA.Offset(1), rngA.Offset(1).Resize(cntR - 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="목록1, 목록2"
        .ErrorTitle = "입력확인"
        .ErrorMessage = "목록에서 선택하여 입력하세요."
    End With
        
End Sub

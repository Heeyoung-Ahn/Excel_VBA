Attribute VB_Name = "sb_AdvancedFilter2"
Option Explicit

'-------------------------------------------
'  xlFilterInPlace의 경우 고급필터 해제
'-------------------------------------------
Sub dismissAF()
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
    End If
End Sub

'------------------------------
'  xlFilterInPlace 고급필터
'------------------------------
Sub AF1()
    Dim rngDB As Range
    Dim rngCriteria As Range
    
    Set rngDB = Sheets("data").Cells(1, 1).CurrentRegion
    Set rngCriteria = Sheets("data").Cells(1, "K").CurrentRegion
    
    rngDB.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngCriteria, Unique:=False
End Sub

'-------------------------------------------------------------
'  복사위치 영역을 정확하게 설정하여 진행하는 고급필터
'    - 복사위치 아래의 데이터를 보존할 수 있음
'-------------------------------------------------------------
Sub AF2()
    Dim rngDB As Range
    Dim rngCriteria As Range
    Dim rngCopy As Range
    Dim i As Integer
    
    With Sheets("data")
        i = Application.WorksheetFunction.CountIfs(.Columns(2), .[k2], .Columns("E:E"), .[l2])
        Set rngDB = .[a1].CurrentRegion
        Set rngCriteria = .[k1].CurrentRegion
        Set rngCopy = .[n1].Resize(i + 1, .[n1].CurrentRegion.Columns.Count)
    End With
    
    rngDB.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:=rngCriteria, copytorange:=rngCopy, Unique:=False
End Sub

'-----------------------
'  일반적인 고급필터
'-----------------------
Sub AF3()
    Dim rngDB As Range
    Dim rngCriteria As Range
    Dim rngCopy As Range
    
    With Sheets("data")
        Set rngDB = .Range("A1").CurrentRegion
        Set rngCriteria = .Range("K1").CurrentRegion
        Set rngCopy = .Range("N1").CurrentRegion.Resize(1)
    End With
    
    rngDB.AdvancedFilter xlFilterCopy, rngCriteria, rngCopy, False '인수의 순서대로 작성 시 인수명 생략 가능
End Sub

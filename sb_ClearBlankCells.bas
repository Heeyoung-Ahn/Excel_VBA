Attribute VB_Name = "sb_ClearBlankCells"
Option Explicit
Const banner As String = "Excel VBA"

'-------------------------------------------------------------
'  수식으로 발생된 빈셀처첨 보이는 유령셀 제거
'  유령셀의 내용을 지우고 찌꺼지 영역을 제거하는 코드
'-------------------------------------------------------------
Sub ClearBlankCells()
On Error GoTo ErrHandler:
    Dim data As Range, SelectedCell As Range, Cell As Range
    Dim cntR As Integer, cntC As Integer
    
    Set SelectedCell = Application.InputBox("데이터가 있는 영역의 아무 셀이나 선택하세요.", banner, Type:=8)
    Set data = SelectedCell.CurrentRegion
    
    '빈셀처럼 보이는 찌거기 데이터가 입력된 셀의 데이터 무두 지우기
    For Each Cell In data
        If Len(Cell) = 0 Then
            Cell.ClearContents
        End If
    Next
    
    '찌꺼기 영역 제거
    cntR = data.Rows.Count
    cntC = data.Columns.Count
    
    data.Cells(cntR + 1, 1).Resize(Rows.Count - cntR, Columns.Count).Delete
    data.Cells(1, cntC + 1).Resize(Rows.Count, Columns.Count - cntC).Delete
    
    ActiveWorkbook.Save
    Exit Sub
    
ErrHandler:
    MsgBox "에러가 발생했습니다." & Space(7) & vbNewLine & _
                  " ※ 에러가 발생한 내용을 캡처하여 관리자에게 보내주세요." & vbNewLine & vbNewLine & _
                  "  ▶ 작업자 : " & Application.UserName & vbNewLine & _
                  "  ▶ 작업일시 : " & Now & vbNewLine & _
                  "  ▶ 에러 코드 : " & Err.Number & vbNewLine & _
                  "  ▶ 에러 내용 : " & Err.Description & vbNewLine & _
                  "  ▶ 에러 소스 : " & Err.Source
End Sub
 

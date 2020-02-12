Attribute VB_Name = "fn_transposeArray"
Option Explicit

'---------------------------------------------------------------------------------
'  이거 쓸바에는 Application.WorksheetFunction.Transpose 함수 이용
'---------------------------------------------------------------------------------
Public Function TransposeArray(InputArr As Variant) As Variant

    Dim RowNdx, ColNdx, LB1, LB2, UB1, UB2 As Long, tmpArray
    
    LB1 = LBound(InputArr, 1)
    LB2 = LBound(InputArr, 2)
    UB1 = UBound(InputArr, 1)
    UB2 = UBound(InputArr, 2)
    
    ReDim tmpArray(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)
    
    For RowNdx = LB2 To UB2
        For ColNdx = LB1 To UB1
            tmpArray(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
        Next ColNdx
    Next RowNdx
    
    TransposeArray = tmpArray

End Function

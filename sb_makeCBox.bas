Attribute VB_Name = "sb_makeCBox"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  단순 콤보 상자 목록 만들기:
'    - makeCBox(콤보박스이름, Array("항목1", "항목2", ...), ListIndex값)
'    - 예: makecCBox(cbo1, Array("전체", "당좌예금", "보통예금", "기타예금", "현금"), -1)
'------------------------------------------------------------------------------------------------------------------
Sub makeCBox(ByRef argCBox As MSForms.ComboBox, ByVal params As Variant, Optional ByVal index As Integer = -1)
    Dim cntParams As Integer
    
    For cntParams = LBound(params) To UBound(params)
        argCBox.AddItem params(cntParams)
    Next cntParams
    
    argCBox.ListIndex = index
End Sub

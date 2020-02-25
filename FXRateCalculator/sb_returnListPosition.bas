Attribute VB_Name = "sb_returnListPosition"
Option Explicit

'------------------------------------------------------------------------------------------------------------------
'  원래의 리스트 항목으로 이동
'    - ReturnListPosition(폼이름, 리스트이름, key값)
'------------------------------------------------------------------------------------------------------------------
Sub returnListPosition(ByRef argForm As UserForm, ByVal argList As String, ByVal argKey As String)
    Dim i As Long
    Dim colKey As Integer 'list에서 key값의 컬럼위치
    
    '//list.BoundColumn의 기본값은 1
    colKey = argForm.Controls(argList).BoundColumn - 1
    
    '//listbox에서 지정된 item위치로 이동
    With argForm.Controls(argList)
        For i = 0 To .ListCount - 1
            If CStr(.List(i, colKey)) = CStr(argKey) Then
               .ListIndex = i
                Exit For
            End If
        Next i
    End With
End Sub


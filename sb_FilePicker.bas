Attribute VB_Name = "sb_FilePicker"
Option Explicit

Sub FilePicker()
    Dim i As Integer
    With Application.FileDialog(msoFileDialogOpen)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        For i = 1 To .SelectedItems.Count
            MsgBox .SelectedItems(i)
        Next i
    End With
End Sub


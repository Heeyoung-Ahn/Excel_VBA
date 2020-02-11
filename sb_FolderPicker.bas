Attribute VB_Name = "sb_FolderPicker"
Option Explicit

Sub FolderPicker()
    Dim rawPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        rawPath = .SelectedItems(1) & Application.PathSeparator
    End With
    MsgBox rawPath
End Sub

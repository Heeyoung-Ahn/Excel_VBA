Attribute VB_Name = "sb_GetOpenfilename"
Option Explicit

'GetOpenfilename(FileFilter, FilterIndex, Title, ButtonText, MultiSelect)
Sub GetOpenFilename()
    Dim fileToOpen As Variant
    Dim eachFile As Variant
    Dim cntArray As Integer
    
    fileToOpen = Application.GetOpenFilename(filefilter:="Excel Files, *.xls*", Title:="Select Excel Files", MultiSelect:=True)
    If TypeName(fileToOpen) = "Boolean" Then Exit Sub
    
    cntArray = UBound(fileToOpen) - LBound(fileToOpen) + 1
    MsgBox cntArray
    
    For Each eachFile In fileToOpen
        MsgBox eachFile
    Next eachFile
End Sub


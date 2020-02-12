Attribute VB_Name = "fn_GetDesktopPath"
Option Explicit

Public Function GetDesktopPath(Optional BackSlash As Boolean = True)

    Dim oWSHShell As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    
    If BackSlash = True Then
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop") & "\"
    Else
        GetDesktopPath = oWSHShell.SpecialFolders("Desktop")
    End If
    
    Set oWSHShell = Nothing

End Function

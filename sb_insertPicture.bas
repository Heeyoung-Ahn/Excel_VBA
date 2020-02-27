Attribute VB_Name = "sb_insertPicture"
Option Explicit

Sub InsertPicture(Optional Pic As Variant = False)
    Dim kk
    Dim Pb
    'Dim Pic As Variant
    
    If Pic = False Then
        Pic = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
    End If
    
    If Pic = False Then
        Exit Sub
    End If
    
    Set Pb = Selection
    ActiveSheet.Pictures.Insert(Pic).Select
    
    kk = WorksheetFunction.Min(Pb.Height / Selection.Height * 0.98, Pb.Width / Selection.Width * 0.98)
    Selection.ShapeRange.ScaleWidth kk, msoFalse, msoScaleFromTopLeft
    With Selection
    .Left = Pb.Left + (Pb.Width - Selection.Width) / 2
    .Top = Pb.Top + (Pb.Height - Selection.Height) / 2
    
    
    End With
End Sub


Sub InsertPicture2(Optional Pic As Variant = False)
    Dim kk
    Dim Pb
    'Dim Pic As Variant
    
    If Pic = False Then
        Pic = Application.GetOpenFilename(fileFilter:="Picturefile, *.jpg;*.jpeg; *.bmp; *.tif; *.png")
    End If
    
    If Pic = False Then
        Exit Sub
    End If
    
    Set Pb = Selection
    ActiveSheet.Pictures.Insert(Pic).Select

    With Selection
        .ShapeRange.LockAspectRatio = msoFalse '그림 좌우고정비율 해제
        .Left = Pb.Left + 2
        .Top = Pb.Top + 2
        .Height = Pb.Height - 2
        .Width = Pb.Width - 4
    End With
End Sub

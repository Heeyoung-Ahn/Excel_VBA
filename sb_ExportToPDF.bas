Attribute VB_Name = "sb_ExportToPDF"
Option Explicit
Public Const banner As String = "PDF파일로 내보내기"

'--------------------------
'  PDF 파일로 내보내기
'--------------------------
Sub exportToPDF()
    '//출력자료 유무 검증 및 인쇄영역 설정
    With Sheet1
        .Activate
        If .[b6] = Empty Then
            MsgBox "작성된 내용이 없습니다.          ", vbInformation, banner
            Exit Sub
        End If
        With .PageSetup
            .PrintArea = Range(Range("B1"), Cells(Rows.Count, "B").End(xlUp).Offset(0, 6)).Address
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait 'xlLandscape(가로방향), xlPortrait(세로방향)
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = 10
        End With
    End With
    
    '//PDF내보내기
    ActiveSheet.ExportAsFixedFormat _
       Type:=xlTypePDF, _
       Filename:=GetDesktopPath() & ActiveSheet.Name & "_" & Range("D8").Value & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".pdf", _
       Quality:=xlQualityStandard, _
       IncludeDocProperties:=True, _
       IgnorePrintAreas:=False, _
       OpenAfterPublish:=True
    
End Sub

'---------------------------------------
'  바탕화면 경로 불러오는 프로시저
'---------------------------------------
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


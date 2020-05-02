Attribute VB_Name = "sb_reportPrint"
Option Explicit
Dim strPrinter As String
Public Const banner As String = "PDF파일로 내보내기"

'------------------
'  프린터 선택
'------------------
Sub Select_Printer()
    Application.Dialogs(xlDialogPrinterSetup).Show
    strPrinter = Application.ActivePrinter
End Sub

'-----------------------------------
'  일일송금보고서 미리보기
'    - Worksheet.PrintPreview
'-----------------------------------
Sub reportPreview()
    With Sheet1
        .Activate
        If .[b6] = Empty Then
            MsgBox "미리 볼 내용이 없습니다.          ", vbInformation, banner
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
        
        .PrintPreview
    End With
End Sub
'------------------------------
'  일일송금보고서 출력
'    - Worksheet.PrintOut
'------------------------------
Sub reportPrint()
    On Error GoTo message
    With Sheet1
        If .[b6] = Empty Then
            MsgBox "인쇄할 내용이 없습니다.                ", vbInformation, banner
            Exit Sub
        End If
        
        '프린터기 설정
        Call Select_Printer
        
        '출력확인
        If MsgBox("보고서를 출력하겠습니까?                     ", vbQuestion + vbYesNo, banner) = vbNo Then Exit Sub
        
        With .PageSetup
            .PrintArea = Range(Range("B1"), Cells(Rows.Count, "B").End(xlUp).Offset(0, 6)).Address
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait 'xlLandscape(가로방향), xlPortrait(세로방향)
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = 10
        End With
        
        Application.ActivePrinter = strPrinter
        .PrintOut
        
    End With
    MsgBox "인쇄가 완료되었습니다.      ", vbInformation, banner
    Exit Sub

message:
  MsgBox "프린터 설정 또는 프린터 오류로 인쇄기능을 사용할 수 없습니다. ", vbInformation, banner
End Sub

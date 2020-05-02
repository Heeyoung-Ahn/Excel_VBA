Attribute VB_Name = "sb_ExportToPDF"
Option Explicit
Public Const banner As String = "PDF���Ϸ� ��������"

'--------------------------
'  PDF ���Ϸ� ��������
'--------------------------
Sub exportToPDF()
    '//����ڷ� ���� ���� �� �μ⿵�� ����
    With Sheet1
        .Activate
        If .[b6] = Empty Then
            MsgBox "�ۼ��� ������ �����ϴ�.          ", vbInformation, banner
            Exit Sub
        End If
        With .PageSetup
            .PrintArea = Range(Range("B1"), Cells(Rows.Count, "B").End(xlUp).Offset(0, 6)).Address
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait 'xlLandscape(���ι���), xlPortrait(���ι���)
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = 10
        End With
    End With
    
    '//PDF��������
    ActiveSheet.ExportAsFixedFormat _
       Type:=xlTypePDF, _
       Filename:=GetDesktopPath() & ActiveSheet.Name & "_" & Range("D8").Value & "(" & Format(Date, "yyyymmdd") & "_" & Format(Time, "hhmm") & ")" & ".pdf", _
       Quality:=xlQualityStandard, _
       IncludeDocProperties:=True, _
       IgnorePrintAreas:=False, _
       OpenAfterPublish:=True
    
End Sub

'---------------------------------------
'  ����ȭ�� ��� �ҷ����� ���ν���
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


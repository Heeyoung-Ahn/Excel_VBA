Attribute VB_Name = "sb_reportPrint"
Option Explicit
Dim strPrinter As String
Public Const banner As String = "PDF���Ϸ� ��������"

'------------------
'  ������ ����
'------------------
Sub Select_Printer()
    Application.Dialogs(xlDialogPrinterSetup).Show
    strPrinter = Application.ActivePrinter
End Sub

'-----------------------------------
'  ���ϼ۱ݺ��� �̸�����
'    - Worksheet.PrintPreview
'-----------------------------------
Sub reportPreview()
    With Sheet1
        .Activate
        If .[b6] = Empty Then
            MsgBox "�̸� �� ������ �����ϴ�.          ", vbInformation, banner
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
        
        .PrintPreview
    End With
End Sub
'------------------------------
'  ���ϼ۱ݺ��� ���
'    - Worksheet.PrintOut
'------------------------------
Sub reportPrint()
    On Error GoTo message
    With Sheet1
        If .[b6] = Empty Then
            MsgBox "�μ��� ������ �����ϴ�.                ", vbInformation, banner
            Exit Sub
        End If
        
        '�����ͱ� ����
        Call Select_Printer
        
        '���Ȯ��
        If MsgBox("������ ����ϰڽ��ϱ�?                     ", vbQuestion + vbYesNo, banner) = vbNo Then Exit Sub
        
        With .PageSetup
            .PrintArea = Range(Range("B1"), Cells(Rows.Count, "B").End(xlUp).Offset(0, 6)).Address
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait 'xlLandscape(���ι���), xlPortrait(���ι���)
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = 10
        End With
        
        Application.ActivePrinter = strPrinter
        .PrintOut
        
    End With
    MsgBox "�μⰡ �Ϸ�Ǿ����ϴ�.      ", vbInformation, banner
    Exit Sub

message:
  MsgBox "������ ���� �Ǵ� ������ ������ �μ����� ����� �� �����ϴ�. ", vbInformation, banner
End Sub

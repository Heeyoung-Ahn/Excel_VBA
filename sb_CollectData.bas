Attribute VB_Name = "sb_CollectData"
Option Explicit
Public Const banner As String = "파일수합프로그램"

'----------------------------------------------
'  폴더에서 여러 파일의 데이터 수합
'    - 기존자료삭제
'    - 폴더선택: FileDialog Property 사용
'    - 폴더 내 파일 유무 검증
'    - 모든 파일 순환하며 자료 수합
'----------------------------------------------
Sub CollectData()

    Dim rawPath As String
    Dim rawFile As String
    Dim taskFile As String
    Dim taskSht As String
    Dim cntFile As Integer
    Dim rngDB As Range
    Dim cntR As Long
    
    Application.ScreenUpdating = False
    
    '//변수설정
    taskFile = ThisWorkbook.Name
    taskSht = "Data" '작업시트이름 ★★
       
    '//기존자료 삭제
    Sheets(taskSht).UsedRange.Offset(1).ClearContents
    
    '//폴더 선택
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        rawPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    '//폴더 내의 엑셀파일을 불러오고, 파일이 없으면 매크로 종료
    rawFile = Dir(rawPath & "*.xls*")
    If rawFile = "" Then
        MsgBox "선택한 폴더에 파일이 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
    '//폴더 내 모든 엑셀파일을 순환
    cntFile = 0
    Do While rawFile <> ""
        Workbooks.Open FileName:=rawPath & rawFile
        Set rngDB = ActiveSheet.Range("A1").CurrentRegion
        rngDB.Offset(1).Resize(rngDB.Rows.Count - 1).Copy
        Workbooks(taskFile).Activate
        Sheets(taskSht).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        Application.DisplayAlerts = False
            Workbooks(rawFile).Close savechanges:=False
        Application.DisplayAlerts = True
        Set rngDB = Nothing
        rawFile = Dir()
        cntFile = cntFile + 1
    Loop
    
    '//찌꺼지 제거
    Set rngDB = Range("A1").CurrentRegion
    cntR = rngDB.Rows.Count
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR).Delete shift:=xlUp
    
    '//마무리
    Application.ScreenUpdating = True
    Range("A1").Activate
    MsgBox cntFile & "개의 파일에서 자료 수합을 완료하였습니다.", vbInformation, banner
End Sub




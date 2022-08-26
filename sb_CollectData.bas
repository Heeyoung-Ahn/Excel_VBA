Attribute VB_Name = "sb_CollectData"
Option Explicit
Public Const banner As String = "파일수합프로그램"

'----------------------------------------------------
'  폴더에서 여러 파일의 데이터 수합
'    - 기존자료삭제
'    - 폴더선택: FileDialog Property 사용
'    - 폴더 내 파일 유무 검증
'    - 모든 파일 순환하며 자료 수합
'    - 파일 구조 검토하여 다르면 Pass하고 알림
'----------------------------------------------------
Sub CollectData()

    '//변수선언
    Dim rawPath As String, rawFile As String, rawSht As String
    Dim taskFile As String, taskSht As String
    Dim taskFieldNM() As Variant, rawFieldNM() As Variant
    Dim cntTC As Integer, cntRC As Integer, cntR As Long, i As Integer
    Dim rngDB As Range
    Dim cntFile As Integer
    
    Application.ScreenUpdating = False
    
    '//변수설정
    taskFile = "파일수합샘플.xlsm"
    taskSht = "Data"
    rawSht = "Sheet1"
        
    '//taskfile 구조 배열에 반환
    Set rngDB = Sheets(taskSht).Range("A1").CurrentRegion.Rows(1)
    cntTC = rngDB.Columns.Count
    ReDim taskFieldNM(1 To cntTC)
    For i = 1 To cntTC
        taskFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
    Next i
       
    '//기존자료 삭제
    Sheets(taskSht).Range("A1").CurrentRegion.Offset(1).ClearContents
        
    '//raw folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        rawPath = .SelectedItems(1) & Application.PathSeparator
    End With
        
    '//rawfile check
    rawFile = Dir(rawPath & "*.xls*")
    If Len(rawFile) = 0 Then
        MsgBox "선택한 폴더에 엑셀 파일이 없습니다.", vbInformation, banner
        Exit Sub
    End If
    
    '//loop
    cntFile = 0
    Do
        Workbooks.Open Filename:=rawPath & rawFile
        Set rngDB = Sheets(rawSht).Range("A1").CurrentRegion.Rows(1)
        cntRC = rngDB.Columns.Count
        'rawfile 구조 배열에 반환
        ReDim rawFieldNM(1 To cntRC)
        For i = 1 To cntRC
            rawFieldNM(i) = rngDB.Cells(1, 1).Offset(0, i - 1).Value
        Next i
        '구조비교1: 필드수
        If cntTC <> cntRC Then
            MsgBox rawFile & "의 필드 수가 TaskFile의 필드 수와 다릅니다." & vbNewLine & _
                    "다음 파일로 건너뜁니다.", vbCritical, banner
                GoTo nextFile:
        End If
        '구조비교2: 필드명
        For i = 1 To cntTC
            If taskFieldNM(i) <> rawFieldNM(i) Then
                MsgBox rawFile & "의 필드명이 TaskFile의 필드명과 다릅니다." & vbNewLine & _
                    "다음 파일로 건너뜁니다.", vbCritical, banner
                GoTo nextFile:
            End If
        Next i
        
        '자료 수합
        rngDB.CurrentRegion.Offset(1).Resize(rngDB.CurrentRegion.Rows.Count - 1).Copy
        Workbooks(taskFile).Activate
        Sheets(taskSht).Cells(Rows.Count, "A").End(xlUp).Offset(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        cntFile = cntFile + 1
        
nextFile:
        '파일정리
        Workbooks(rawFile).Close savechanges:=False
        
        '다음파일
        rawFile = Dir()
    Loop Until rawFile = ""
    
    '//찌꺼기 제거
    Set rngDB = Range("A1").CurrentRegion
    cntR = rngDB.Rows.Count
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
    
    Application.ScreenUpdating = True
    
    '//작업보고
    Range("A1").Activate
    MsgBox cntFile & "개의 파일 수합 완료", vbInformation
End Sub

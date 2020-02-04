Attribute VB_Name = "sb_DirSample"
Option Explicit

'---------------------------------
'  Dir 함수 샘플
'    - 특정 파일 존재 여부 확인
'---------------------------------
Sub DirSample1()
    Dim strFile As String
    strFile = Dir("C:\00 공통기초자료\*교회목록*.xlsx")
    If Len(strFile) = 0 Then
        MsgBox "찾는 파일이 존재하지 않습니다.", vbCritical, "파일찾기"
    Else
        MsgBox "찾는 파일의 이름은 '" & strFile & "'입니다."
    End If
End Sub

'-------------------------------
'  Dir 함수 샘플
'    - 폴더 내 엑셀 파일 찾기
'    - 엑셀 파일 갯수 출력
'    - 엑셀 파일 이름 출력
'-------------------------------
Sub DirSample2()
    Dim strAPath As String
    Dim strAFile As String
    Dim strFile As String
    Dim strFileSet As String
    Dim cntFile As Integer
    
    strAFile = ActiveWorkbook.FullName
    strAPath = Left(strAFile, InStrRev(strAFile, Application.PathSeparator))
    strFile = Dir(strAPath & "*.xls*")
    
    '//파일 갯수 확인 및 출력
    Do While strFile <> ""
        cntFile = cntFile + 1
        strFile = Dir
    Loop
    MsgBox "'" & strAPath & "' 폴더 내 엑셀 파일의 갯수는 " & cntFile & "개입니다.", vbInformation, "파일갯수조회"
    
    '//파일명 출력
    strFile = Dir(strAPath & "*.xls*")
    Do
        strFileSet = strFileSet & strFile & vbNewLine
        strFile = Dir
    Loop Until strFile = ""
    MsgBox strFileSet, vbInformation, "엑셀파일 이름 조회"
End Sub

'------------------------------------------------------------------
'  Dir 함수 샘플
'    - FileDialog Property 속성 이용 폴더 선택
'    - 하위 폴더 반환
'    - Msgbox에 출력
'------------------------------------------------------------------
Sub DirSample3()
    Dim strAPath As String
    Dim strAFile As String
    Dim strSubPath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        strAPath = .SelectedItems(1) & Application.PathSeparator
    End With
    
    strSubPath = Dir(strAPath, vbDirectory)
    Do While strSubPath <> ""
        If strSubPath <> "." And strSubPath <> ".." Then
            If (GetAttr(strAPath & strSubPath) And vbDirectory) = vbDirectory Then
                MsgBox strSubPath
            End If
        End If
        strSubPath = Dir
    Loop
End Sub


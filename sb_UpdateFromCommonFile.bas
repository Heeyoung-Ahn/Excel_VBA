Attribute VB_Name = "sb_UpdateFromCommonFile"
Option Explicit
Const banner As String = "공통기초자료업데이트"

'---------------------------------------------------------------------
'  공통폴더의 공통기초자료 파일을 열어서 작업 파일 업데이트
'---------------------------------------------------------------------
Sub UpdateFromCommonFile()

On Error Resume Next
    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String, tskS As String
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim DB As Range
    Dim rawFOpen As Boolean

    '//변수 정의
    rawF = "원본파일이름" '★★
    rawS = "원본시트이름" '★★
    tskF = ThisWorkbook.Name '작업파일 이름
    tskS = "작업시트이름" '★★
       
    '//공통DB 폴더를 돌면서 업데이트 대상 파일 찾아서 rawF에 이름 설정
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\01 공통DB\" '공통 폴더 경로 설정★★
        rawF = Dir(rawP & rawF) '원본파일 이름 경로 포함 설정
        If Left(rawF, 1) = "~" Then '파일을 다른 사람이 열고 있는 경우
            MsgBox "공통기초자료 파일이 다른 누군가에 의해 열려 있습니다." & vbNewLine & _
                "확인바랍니다." & Space(10), vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox "공통기초자료 파일이 공통DB 폴더에 없습니다." & vbNewLine & _
        "확인바랍니다." & Space(10), vbInformation, banner
    Exit Sub
n:

    '//업데이트 대상 파일 열기
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then rawFOpen = True
        Exit For
    Next
    If rawFOpen = True Then
        Windows(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF, Password:="파일의 비밀번호"   '비밀번호로 파일 열기★★
        Windows(rawF).Activate
    End If
    
    '//작업파일의 기초자료 초기화
    Windows(tskF).Activate
    Sheets(tskS).UsedRange.ClearContents
    
    '//공통기초자료에서 기초자료 가져오기
    Windows(rawF).Activate
    Sheets(rawS).UsedRange.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '//업데이트 대상 파일 닫기
    Windows(rawF).Close savechanges:=False '저장안하고 닫기
       
    '//데이터영역설정
    Windows(tskF).Activate
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '//찌꺼기 영역 삭제
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete Shift:=xlUp
            
    '//열너비조정
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    
    '//2행기준 서식적용
    Rows("2:2").Copy
    Rows("3:" & cntR).PasteSpecial (4)
    Application.CutCopyMode = False
        
    '//마무리
    ActiveWorkbook.Save
    
End Sub


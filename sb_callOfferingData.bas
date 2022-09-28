Attribute VB_Name = "sb_callOfferingData"
Option Explicit

Dim MName As String
Dim tskResultCD As Integer

Sub check_update()
    MName = "전세계 봉헌 데이터" '업데이트할 파일이름 설정 ★★

    If MsgBox(MName & " 자료를 공통기초자료 폴더에서 업데이트합니다.           " & Chr(13) & _
        "준비되었습니까?                ", vbQuestion + vbYesNo, banner) = vbNo Then
        MsgBox "그럼 다시 준비하고 업데이트를 진행해 주세요.         ", vbInformation, banner
        Exit Sub
    End If
    
    Call OfferingDBUpdate '작업 프로시저 설정 ★★
    
    Sheets("작업").Activate
    
    If tskResultCD = 1 Then
        MsgBox MName & " 자료 업데이트가 완료되었습니다.    ", vbInformation, banner
    End If

End Sub

Sub OfferingDBUpdate()

    Dim fileC As Workbook
    Dim rawP As String, rawF As String, rawS As String
    Dim tskF As String, tskS As String
    Dim DB As Range
    Dim cntR As Integer, cntC As Integer, i As Integer
    Dim rawFOpen As Boolean
    Dim oldFieldNM() As String, newFieldNM() As String

    '매크로 최적화
    Call Optimization

    '파일 전체 변수 정의
    tskF = ThisWorkbook.Name '작업파일 이름 설정
    
    '공통기초자료 폴더에서 업데이트 대상 파일을 찾아서 rawF에 설정
    For i = 1 To 24
        rawP = Chr(66 + i) & ":\00 공통기초자료\" '업데이트 대상 자료의 폴더 설정 ★★
        rawF = Dir(rawP & "*20 전세계 봉헌금 데이터*") '원본파일 이름 설정 ★★
        If Left(rawF, 1) = "~" Then
            MsgBox MName & " 파일을 다른 누군가가 열고 있습니다.   " & vbCrLf & _
                "확인 후 다시 진행해 주세요.            ", vbInformation, banner
            Exit Sub
        End If
        If rawF <> Empty Then GoTo n:
    Next
    MsgBox MName & " 파일이 업데이트하려는 폴더에 없습니다.   " & vbCrLf & _
        "확인 후 다시 진행해 주세요.                ", vbInformation, banner
    Exit Sub
n:
    
    '어드민 파일 열기
    rawFOpen = False
    For Each fileC In Workbooks
        If fileC.Name = rawF Then
            rawFOpen = True
            Exit For
        End If
    Next
    If rawFOpen = True Then
        Windows(rawF).Activate
    Else
        Workbooks.Open Filename:=rawP & rawF ', Password:="qaz1234" '비밀번호로 파일 열기★★
        Windows(rawF).Activate
    End If
    
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '첫 번째 시트 변수 정의
    rawS = "지교회 회계 관리(독립채산제)를 위한 교회별 봉헌내역" '원본시트 이름 설정 ★★
    tskS = "t_church_offering_yyyymm_temp" '작업시트 이름 설정 ★★
    
    '기존 작업파일 필드명 oldFieldNM 배열에 반환
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '어드민의 병합된 필드명 제거
    'Rows("1:1").Delete shift:=xlUp '1행에 병합된 필드명 삭제 필요시 진행★★
        
    '어드민 파일 필드명 newFieldNM 배열에 반환
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '파일 구조 점검: 필드명
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " 파일의 " & tskS & " 시트의 어드민 자료와 엑셀자료의 필드명이 서로 불일치합니다." & Chr(13) & _
                "확인 후 다시 진행해 주세요.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo m:
        End If
    Next i
            
    '기초자료 초기화
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '어드민 자료 가져오기
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '데이터영역설정
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '찌꺼기 영역 삭제
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '서식정리
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0값만 있는 경우 지우기
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '필요시 진행 ★★

    '열너비조정
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

m:

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '두 번째 시트 변수 정의
    rawS = "지교회별 봉헌자수 정보" '원본시트 이름 설정 ★★
    tskS = "t_church_offering_saint_no_yyyy" '작업시트 이름 설정 ★★
    
    '기존 작업파일 필드명 oldFieldNM 배열에 반환
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '어드민의 병합된 필드명 제거
    'Rows("1:1").Delete shift:=xlUp '1행에 병합된 필드명 삭제 필요시 진행★★
        
    '어드민 파일 필드명 newFieldNM 배열에 반환
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '파일 구조 점검: 필드명
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " 파일의 " & tskS & " 시트의 어드민 자료와 엑셀자료의 필드명이 서로 불일치합니다." & Chr(13) & _
                "확인 후 다시 진행해 주세요.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo k:
        End If
    Next i
            
    '기초자료 초기화
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '어드민 자료 가져오기
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '데이터영역설정
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '찌꺼기 영역 삭제
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '서식정리
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0값만 있는 경우 지우기
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '필요시 진행 ★★

    '열너비조정
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

k:

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '세 번째 시트 변수 정의
    rawS = "교회리스트" '원본시트 이름 설정 ★★
    tskS = "t_church_disp_key_info_temp" '작업시트 이름 설정 ★★
    
    '기존 작업파일 필드명 oldFieldNM 배열에 반환
    Windows(tskF).Activate
    Sheets(tskS).Activate
    cntC = Range("A1").CurrentRegion.Columns.Count
    ReDim oldFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        oldFieldNM(i) = Sheets(tskS).Range("A1").Offset(0, i).Value
    Next i

    '어드민의 병합된 필드명 제거
    'Rows("1:1").Delete shift:=xlUp '1행에 병합된 필드명 삭제 필요시 진행★★
        
    '어드민 파일 필드명 newFieldNM 배열에 반환
    Windows(rawF).Activate
    Sheets(rawS).Activate
    ReDim newFieldNM(cntC - 1)
    For i = 0 To cntC - 1
        newFieldNM(i) = Sheets(rawS).Range("A1").Offset(0, i).Value
    Next i
       
    '파일 구조 점검: 필드명
    For i = 0 To cntC - 1
        If oldFieldNM(i) <> newFieldNM(i) Then
            MsgBox MName & " 파일의 " & tskS & " 시트의 어드민 자료와 엑셀자료의 필드명이 서로 불일치합니다." & Chr(13) & _
                "확인 후 다시 진행해 주세요.   ", vbInformation, banner
            tskResultCD = 0
            Windows(tskF).Activate
            GoTo s:
        End If
    Next i
            
    '기초자료 초기화
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").CurrentRegion.ClearContents
        
    '어드민 자료 가져오기
    Windows(rawF).Activate
    Sheets(rawS).[a1].CurrentRegion.Copy
    Windows(tskF).Activate
    Sheets(tskS).Range("A1").PasteSpecial (3)
    Application.CutCopyMode = False
    
    '데이터영역설정
    Set DB = Sheets(tskS).Range("A1").CurrentRegion
    cntR = DB.Rows.Count
    cntC = DB.Columns.Count
    
    '찌꺼기 영역 삭제
    Sheets(tskS).Activate
    Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(Rows.Count - cntR, Columns.Count).Delete shift:=xlUp
      
    '서식정리
    Range("2:2").Copy
    Range("2:2").Resize(Cells(Rows.Count, 1).End(xlUp).Row - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    '0값만 있는 경우 지우기
    'DB.Replace what:="0", replacement:="", lookat:=xlWhole '필요시 진행 ★★

    '열너비조정
    Sheets(tskS).UsedRange.EntireColumn.AutoFit
    tskResultCD = 1

s:

    '어드민 파일이 닫혀있었다면 다시 닫기
    If rawFOpen = False Then
        Windows(rawF).Activate
        Windows(rawF).Close SaveChanges:=False
    End If

    '마무리
    ActiveWorkbook.Save

    '매크로 최적화 원복
    Call Normal
    
End Sub






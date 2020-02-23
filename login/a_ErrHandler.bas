Attribute VB_Name = "a_ErrHandler"
Option Explicit

'-----------------------------------------------------------------------------------------------------
'  에러처리: errhandle(프로시저명, 테이블명, SQL문, 폼이름, 작업명)
'    - 에러 발생 내용을 디버깅 하기 위해 메시지 박스로 표시
'    - 에러 발생에 따른 로그 기록은 DB와 관계된 내용만 callDBroRS, executeSQL에서 진행
'-----------------------------------------------------------------------------------------------------
Sub ErrHandle(ProcedureNM As String, Optional tableNM As String = "NULL", Optional SQLScript As String = "NULL", Optional formNM As String = "NULL", Optional JobNM As String = "기타")
    If Err.Number <> 0 Then
        MsgBox "에러가 발생했습니다." & Space(7) & vbNewLine & _
            " ※ 에러가 발생한 내용을 캡처하여 관리자에게 보내주세요." & vbNewLine & vbNewLine & _
            "  ▶ 작업자 : " & Application.UserName & vbNewLine & _
            "  ▶ 작업일시 : " & Now & vbNewLine & _
            "  ▶ 작업내용 : " & JobNM & vbNewLine & vbNewLine & _
            "  ▶ 오류 발생 vba : " & ProcedureNM & vbNewLine & _
            "  ▶ 오류 발생 폼 : " & formNM & vbNewLine & _
            "  ▶ 오류 발생 DB : " & tableNM & vbNewLine & _
            "  ▶ 오류 발생 Script : " & SQLScript & vbNewLine & vbNewLine & vbNewLine & _
            "  ▶ 에러 코드 : " & Err.Number & vbNewLine & _
            "  ▶ 에러 내용 : " & Err.Description & vbNewLine & _
            "  ▶ 에러 소스 : " & Err.Source
    End If
End Sub


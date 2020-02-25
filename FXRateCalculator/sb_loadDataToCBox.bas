Attribute VB_Name = "sb_loadDataToCBox"
Option Explicit

'----------------------------------------------------------------
'  콤보박스 리스팅
'    - loadDataToCBox(콤보박스, SQL문, DB)
'----------------------------------------------------------------
Sub loadDataToCBox(argCboBox As MSForms.ComboBox, argSQL As String, argDB As String, argFormNM As String)
    Dim i As Integer, j As Integer
    Dim listData() As String
    
    Call connectTaskDB
    callDBtoRS "loadDataToCBox", argDB, argSQL, argFormNM, "콤보박스리스팅"
    'Debug.Print argSQL
    '//레코드셋의 데이터를 listData 배열에 반환
    If rs.EOF Then
        'MsgBox argFormNM & "의 " & argCboBox.Name & "에 구성할 자료가 없습니다.", vbInformation, Banner
        argCboBox.Clear
        disconnectALL
        Exit Sub
    End If
    ReDim listData(0 To rs.RecordCount - 1, 0 To rs.Fields.Count - 1) '//DB에서 반환할 배열의 크기 지정: 레코드셋의 레코드 수, 필드 수
    rs.MoveFirst
    For i = 0 To rs.RecordCount - 1
        For j = 0 To rs.Fields.Count - 1
            listData(i, j) = rs.Fields(j).Value
        Next j
        rs.MoveNext
    Next i
    Call disconnectALL
    
    '//listData 배열로 반환된 Data를 콤보박스에 리스팅
    argCboBox.List = listData
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_calendar 
   Caption         =   "Calendar"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3330
   OleObjectBlob   =   "f_calendar.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "f_Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim CYear As Integer, CMonth As Integer, WeekNum As Integer
Dim DateV As Date
Dim Str As String
Const s_year As Integer = 2000 '조회시작연도
Const e_year As Integer = 2030 '조회마지막연도

Private Sub UserForm_Initialize()
    For i = s_year To e_year '연도 리스트 생성
        cbo_year.AddItem i
    Next i
    cbo_year.ListIndex = Year(Date) - s_year
    For i = 1 To 12 '월 리스트 생성
        cbo_month.AddItem i
    Next i
    cbo_month.ListIndex = Month(Date) - 1
    Cal
End Sub

Function Cal()
    CYear = cbo_year.ListIndex + s_year '연도
    CMonth = cbo_month.ListIndex + 1 '월
    WeekNum = Weekday(DateSerial(CYear, CMonth, 1)) '1일 시작요일

    For i = 0 To 41 '총 42개 버튼에 날짜 입력 및 글꼴 설정
        DateV = DateSerial(CYear, CMonth, 1 + (7 - WeekNum) - 6 + i)
        Str = "cmd_day_" & i + 1
        
        If Year(DateV) = CYear And Month(DateV) = CMonth Then
            With Me.Controls(Str)
                .Caption = Day(DateV)
                .Enabled = True
                .Font.Bold = False
            End With
        Else
            With Me.Controls(Str) '이번달의 날짜가 아닐경우 캡션 미표시 및 비활성화
                .Caption = ""
                .Enabled = False
                .Font.Bold = False
            End With
        End If
        
    If Year(DateV) = Year(Date) And Month(DateV) = Month(Date) And Day(DateV) = Day(Date) Then '오늘 날짜일 경우 굵게 및 포커스 활성화
        With Me.Controls(Str)
            If .Enabled = True Then
                .SetFocus
                .Font.Bold = True
            End If
        End With
    End If
    Next i
End Function

Private Sub cbo_month_Change()
    Cal
End Sub
Private Sub cbo_year_Change()
    Cal
End Sub

Private Sub cmd_day_1_Click()
    Day_Click
End Sub
Private Sub cmd_day_2_Click()
    Day_Click
End Sub
Private Sub cmd_day_3_Click()
    Day_Click
End Sub
Private Sub cmd_day_4_Click()
    Day_Click
End Sub
Private Sub cmd_day_5_Click()
    Day_Click
End Sub
Private Sub cmd_day_6_Click()
    Day_Click
End Sub
Private Sub cmd_day_7_Click()
    Day_Click
End Sub
Private Sub cmd_day_8_Click()
    Day_Click
End Sub
Private Sub cmd_day_9_Click()
    Day_Click
End Sub
Private Sub cmd_day_10_Click()
    Day_Click
End Sub
Private Sub cmd_day_11_Click()
    Day_Click
End Sub
Private Sub cmd_day_12_Click()
    Day_Click
End Sub
Private Sub cmd_day_13_Click()
    Day_Click
End Sub
Private Sub cmd_day_14_Click()
    Day_Click
End Sub
Private Sub cmd_day_15_Click()
    Day_Click
End Sub
Private Sub cmd_day_16_Click()
    Day_Click
End Sub
Private Sub cmd_day_17_Click()
    Day_Click
End Sub
Private Sub cmd_day_18_Click()
    Day_Click
End Sub
Private Sub cmd_day_19_Click()
    Day_Click
End Sub
Private Sub cmd_day_20_Click()
    Day_Click
End Sub
Private Sub cmd_day_21_Click()
    Day_Click
End Sub
Private Sub cmd_day_22_Click()
    Day_Click
End Sub
Private Sub cmd_day_23_Click()
    Day_Click
End Sub
Private Sub cmd_day_24_Click()
    Day_Click
End Sub
Private Sub cmd_day_25_Click()
    Day_Click
End Sub
Private Sub cmd_day_26_Click()
    Day_Click
End Sub
Private Sub cmd_day_27_Click()
    Day_Click
End Sub
Private Sub cmd_day_28_Click()
    Day_Click
End Sub
Private Sub cmd_day_29_Click()
    Day_Click
End Sub
Private Sub cmd_day_30_Click()
    Day_Click
End Sub
Private Sub cmd_day_31_Click()
    Day_Click
End Sub
Private Sub cmd_day_32_Click()
    Day_Click
End Sub
Private Sub cmd_day_33_Click()
    Day_Click
End Sub
Private Sub cmd_day_34_Click()
    Day_Click
End Sub
Private Sub cmd_day_35_Click()
    Day_Click
End Sub
Private Sub cmd_day_36_Click()
    Day_Click
End Sub
Private Sub cmd_day_37_Click()
    Day_Click
End Sub
Private Sub cmd_day_38_Click()
    Day_Click
End Sub
Private Sub cmd_day_39_Click()
    Day_Click
End Sub
Private Sub cmd_day_40_Click()
    Day_Click
End Sub
Private Sub cmd_day_41_Click()
    Day_Click
End Sub
Private Sub cmd_day_42_Click()
    Day_Click
End Sub

Private Sub cmd_Today_Click()
    cbo_year.ListIndex = Year(Date) - s_year
    cbo_month.ListIndex = Month(Date) - 1
    Str = "cmd_day_" & Day(Date) + WeekNum - 1
    Me.Controls(Str).SetFocus
    Day_Click
End Sub

Private Sub cmd_nextM_Click()
    If cbo_month.ListIndex = 11 Then
        If cbo_year.ListIndex = cbo_year.ListCount - 1 Then
            Exit Sub
        Else
            cbo_year.ListIndex = cbo_year.ListIndex + 1
            cbo_month.ListIndex = 0
        End If
    Else
        cbo_month.ListIndex = cbo_month.ListIndex + 1
    End If
End Sub
Private Sub cmd_preM_Click()
    If cbo_month.ListIndex = 0 Then
        If cbo_year.ListIndex = 0 Then
            Exit Sub
        Else
            cbo_year.ListIndex = cbo_year.ListIndex - 1
            cbo_month.ListIndex = 11
        End If
    Else
        cbo_month.ListIndex = cbo_month.ListIndex - 1
    End If
End Sub

Function Day_Click()
    On Error Resume Next
    MsgBox Format(DateSerial(CYear, CMonth, ActiveControl.Caption), "yyyy-mm-dd")
'    Select Case Calint
'        Case 1
'            FX_C.txt5 = Format(DateSerial(CYear, CMonth, ActiveControl.Caption), "yyyy-mm-dd")
'        Case 2
'            FX_C.txtC1 = Format(DateSerial(CYear, CMonth, ActiveControl.Caption), "yyyy-mm-dd")
'    End Select
    Unload Me
End Function

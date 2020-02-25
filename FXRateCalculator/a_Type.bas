Attribute VB_Name = "A_Type"
Option Explicit

'//구조체 정의
Type t_users '//common.users
    user_id As Integer
    user_gb As String
    user_nm As String
    user_pw As String
    argIP As String
    argDB As String
    argUN As String
    argPW As String
    suspended As Integer '1: suspended
End Type

Type t_currency_cal '//co_account.currency_cal
    currency_id As Integer
    currency_un As String
    refer_dt As Date
    fx_rate_krw As Currency
    fx_rate_usd As Currency
    user_id As Integer
End Type
 
Type t_result
    strSQL As String
    affectedCount As Long
End Type


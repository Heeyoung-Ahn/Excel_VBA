Attribute VB_Name = "sb_setCBox"
Option Explicit

'--------------------------------------------------------------
'  ÄÞº¸¹Ú½º ¼³Á¤: ÀüÃ¼¸ñ·Ï
'    - setCBox(ÄÞº¸¹Ú½ºÀÌ¸§, ÄÞº¸¹Ú½º³»¿ë ¾à¾î, ÆûÀÌ¸§)
'--------------------------------------------------------------
Sub setCBox(ByRef argCBox As MSForms.ComboBox, kindCBox As String, formNM As String, Optional referDate As Date, Optional user_authority As Integer = 1)
    Dim strSQL As String
    If referDate = Empty Then referDate = today
    Select Case kindCBox
    
        Case "FX" '//È­Æó ÄÞº¸
            With argCBox
                .ColumnCount = 3
                .ColumnHeads = False
                .ColumnWidths = "0,50,100" 'È­Æóid, È­Æó¾àÄª, È­Æó¸íÄª
                .TextColumn = 2
                .ListWidth = "150"
                .TextAlign = fmTextAlignLeft
                .IMEMode = fmIMEModeAlpha
                .Style = fmStyleDropDownCombo
            End With
            strSQL = "SELECT currency_id, currency_un, currency_nm FROM co_account.v_currencies ORDER BY sort_order;"
            loadDataToCBox argCBox, strSQL, "co_account.v_currencies", formNM
            
    End Select
End Sub

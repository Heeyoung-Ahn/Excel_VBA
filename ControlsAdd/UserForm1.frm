VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mycmd As Control

Private Sub CommandButton1_Click()
 
    Set Mycmd = Me.Controls.Add("Forms.commandbutton.1")
    With Mycmd
        .Left = 18
        .Top = 150
        .Width = 175
        .Height = 20
        .Name = "cmd_test"
        .Caption = "This is fun." & Mycmd.Name
    End With
 
End Sub
 
Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)
    Label1.Caption = "Control was Added."
End Sub



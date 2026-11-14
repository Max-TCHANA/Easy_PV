VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserAbout 
   Caption         =   "About E@sy_PV"
   ClientHeight    =   5265
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   4380
   OleObjectBlob   =   "UserAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("TEST").Range("M68").Value = 0
Worksheets("GUIDE").Visible = True
With ActiveWindow
.DisplayHorizontalScrollBar = True ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = True ' SUPPRIME : Ascenseur vertical
End With
Worksheets("ACCUEIL").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
Unload UserAbout
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton2_Click()
Unload UserAbout
End Sub

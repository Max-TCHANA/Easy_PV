VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserCastype 
   Caption         =   "Modèles prédéfinis  types:  habitat"
   ClientHeight    =   7275
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   8160
   OleObjectBlob   =   "UserCastype.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserCastype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH0
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton10_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH9
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton11_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH5
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton12_Click()
UserPredefini_PREC
End Sub

Private Sub CommandButton13_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH10
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH1
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton3_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH2
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton4_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH3
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton5_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH4
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton7_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH6
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton8_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH7
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton9_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Unload UserCastype
Worksheets("RECAPITUL").Visible = True
Worksheets("PREDEF").Visible = False
Worksheets("RECAPITUL").Unprotect Password:="max"
MPTH8
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub UserForm_Initialize()
Label1.Caption = UCase("projet " & Worksheets("TEST").Range("R6"))
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then 'Si on clique sur la croix
        Cancel = 1 'Invalide la fermeture
End If
End Sub

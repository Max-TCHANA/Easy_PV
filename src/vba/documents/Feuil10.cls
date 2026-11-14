VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
RECAPITUL_PREC
End Sub

Private Sub CommandButton2_Click()
RECAPITUL_SUIV
End Sub

Private Sub CommandButton3_Click()
imprim_RECAPIT
End Sub

Private Sub CommandButton4_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 3
Worksheets("GUIDE").Visible = True
Worksheets("RECAPITUL").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

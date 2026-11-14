VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
MsgBox "Entrer les informations relatives à votre consommation journalière" & Chr(10) & "Seules les données valides sont prises en compte." & Chr(10) & _
"Ce mode est commode lorsque le mode guidé présente des manquements par rapport à vos besoins.", vbOKOnly + vbInformation, "Mode libre"
End Sub

Private Sub CommandButton2_Click()
DATACH2_PREC
End Sub

Private Sub CommandButton3_Click()
DATACH2_SUIV
End Sub

Private Sub CommandButton4_Click()
efface_DATACH2
End Sub

Private Sub CommandButton5_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 2
Worksheets("GUIDE").Visible = True
Worksheets("DATACH2").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

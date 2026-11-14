VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
efface_DATACH1
End Sub

Private Sub CommandButton10_Click()
DATACH1_SUIV
End Sub

Private Sub CommandButton11_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 1
Worksheets("GUIDE").Visible = True
Worksheets("DATACH1").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

Private Sub CommandButton2_Click()
DATACH1_PREC
End Sub

Private Sub CommandButton3_Click()
DATACH1_SUIV
End Sub

Private Sub CommandButton4_Click()
MsgBox "Ce mode vous permet d'évaluer facilement votre consommation journalière" & Chr(10) & "en vous présentant une variété d'équipements par classe d'utilisation domestique." & Chr(10) & _
"Vous avez juste à indiquer le nombre d'unité et la durée d'utilisation journalière" & Chr(10) & _
" à droite de l'appareil qui vous intéresse." & Chr(10) & _
"Seules les données valides sont prises en compte", vbOKOnly + vbInformation, "Mode guidé"
End Sub

Private Sub CommandButton5_Click()
efface_DATACH1
End Sub

Private Sub CommandButton6_Click()
DATACH1_PREC
End Sub

Private Sub CommandButton7_Click()
DATACH1_SUIV
End Sub

Private Sub CommandButton8_Click()
efface_DATACH1
End Sub

Private Sub CommandButton9_Click()
DATACH1_PREC
End Sub

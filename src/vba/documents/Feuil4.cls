VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
imprim_EQPR
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ACCUEIL").Visible = True
With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
Worksheets("EQPR").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton3_Click()
EQPR_PREC
Application.ScreenUpdating = True
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
End Sub

Private Sub CommandButton4_Click()
Worksheets("EQPR").Range("B88").Select
End Sub

Private Sub CommandButton5_Click()
Worksheets("EQPR").Range("B7").Select
End Sub

Private Sub CommandButton6_Click()
Load UserEquipement
UserEquipement.Show
End Sub

Private Sub CommandButton7_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 6
Worksheets("GUIDE").Visible = True
Worksheets("EQPR").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

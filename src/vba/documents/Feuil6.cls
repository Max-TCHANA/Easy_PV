VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
imprim_PREDIM
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ETATPRO").Visible = True
Worksheets("PREDIM").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton3_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ACCUEIL").Visible = True
With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
Worksheets("PREDIM").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton4_Click()
Courantmax = 1.5 * (Worksheets("PREDIM").Range("C28").Value)
If Courantmax > 1410 Then
MsgBox "Le courant maximal de charge est supérieur à 1410A" & Chr(10) & "Il existe pas de section de câble adaptée a un tel courant" & Chr(10) & _
"Par conséquent, E@sy_PV ne peut proposer des équipements.", vbOKOnly + vbExclamation, "Valeur limite atteinte!!!"
Else
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("EQPR").Visible = True
Worksheets("PREDIM").Visible = False
Application.ScreenUpdating = True
Worksheets("EQPR").Range("A1").Select
Worksheets("EQPR").protect Password:="max"
Worksheets("EQPR").EnableSelection = xlNoSelection
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub

Private Sub CommandButton5_Click()
 PREDIM_PREC
End Sub

Private Sub CommandButton6_Click()
Worksheets("PREDIM").Range("E60").Select
End Sub

Private Sub CommandButton7_Click()
Worksheets("PREDIM").Range("B5").Select
End Sub

Private Sub CommandButton8_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 4
Worksheets("GUIDE").Visible = True
Worksheets("PREDIM").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

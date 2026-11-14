VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'Récupération du choix
choix = Worksheets("TEST").Range("J17").Value
If choix = 1 Then
MsgBox "Comme vous pouvez le voir" & Chr(10) & "Il n'y a pas d'information à imprimer" & Chr(10) & _
"Votre site est à climat mal connu.", vbOKOnly + vbInformation, "Site au Climat mal connu"
Else
imprim_ETATPRO
End If
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("PREDIM").Visible = True
Worksheets("ETATPRO").Visible = False
Application.ScreenUpdating = True
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton3_Click()

ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'valeur sentinelle
Worksheets("TEST").Range("M68").Value = 5
Worksheets("GUIDE").Visible = True
Worksheets("ETATPRO").Visible = False
Worksheets("GUIDE").Select
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub

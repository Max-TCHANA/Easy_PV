Attribute VB_Name = "ACTIONS"
Sub Usage_info()
Attribute Usage_info.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Usage_info Macro
'

'
    Range("AA10").Select
End Sub
Sub Usage_Elairage()
Attribute Usage_Elairage.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Usage_Elairage Macro
'

'
    Range("C10").Select
End Sub
Sub Electro()
Attribute Electro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Electro Macro
'

'
    Range("R12").Select
End Sub
Sub AUdio()
Attribute AUdio.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AUdio Macro
'

'
    Range("X13").Select
End Sub
Sub Réfré()
Attribute Réfré.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Réfré Macro
'

'
    Range("K13").Select
End Sub

Sub Mask()
Set F3 = Worksheets("Feuil3")
F3.Range("A1") = "ça marche!!!"
End Sub
Sub Masque()
Attribute Masque.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Masque Macro
'
Application.ScreenUpdating = False
Sheets("Feuil2").Visible = False
'
'    Sheets("Feuil2").Select
'    ActiveWindow.SelectedSheets.Visible = False
    Application.ScreenUpdating = True
End Sub
Sub Affiche()
Attribute Affiche.VB_ProcData.VB_Invoke_Func = " \n14"
'
Application.ScreenUpdating = False
'

'
'    Sheets("Feuil3").Select
    Sheets("Feuil2").Visible = True
    Sheets("Feuil2").Select
    Application.ScreenUpdating = True
End Sub

Sub Userbeg()
Set F = Worksheets("ACCUEIL")
Set F1 = Worksheets("TEST")
F.Select
F1.Range("R6").Value = ""
Load USERBEGIN
USERBEGIN.Show
End Sub

Sub DATACH1_SUIV()
ActiveWorkbook.Unprotect Password:="max"
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Consommation_guide
'On s'assure que l'utilisateur remplit correctementla feuille
If F1.Range("E4").Value = 0 Then
MsgBox "La consommation journalière est nulle." & Chr(10) & "Veuillez vérifier vos données SVP" & Chr(10) & _
"Vous devez choisir de manière valide au moins 1 élément pour continuer.", vbOKOnly + vbInformation, "Besoins énergétiques nuls"
'Set F1 = worksheets("TEST")
Else
Application.ScreenUpdating = False
Worksheets("RECAPITUL").Visible = True
Worksheets("DATACH1").Visible = False
F.Select
Worksheets("RECAPITUL").Unprotect Password:="max"
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
efface_RECAPITUL
RECAPITULE
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub
Sub DATACH2_SUIV()
ActiveWorkbook.Unprotect Password:="max"
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
consommation_libre
'On s'assure que l'utilisateur remplit correctementla feuille
If F1.Range("E4").Value = 0 Then
MsgBox "La consommation journalière est nulle." & Chr(10) & "Veuillez vérifier vos données SVP" & Chr(10) & _
"Vous devez indiquer au moins 1 élément valide pour continuer.", vbOKOnly + vbInformation, "Besoins énergétiques nuls"
'Set F1 = worksheets("TEST")
Else
Application.ScreenUpdating = False
Worksheets("RECAPITUL").Visible = True
Worksheets("DATACH2").Visible = False
F.Select
Worksheets("RECAPITUL").Unprotect Password:="max"
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
efface_RECAPITUL
RECAPITULE
Worksheets("RECAPITUL").protect Password:="max"
Worksheets("RECAPITUL").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub
Sub DATACH1_PREC()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ACCUEIL").Visible = True
With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
Worksheets("DATACH1").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub DATACH2_PREC()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ACCUEIL").Visible = True
With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
Worksheets("DATACH2").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub RECAPITUL_PREC()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
If Worksheets("TEST").Range("R7").Value = 1 Then
Worksheets("DATACH1").Visible = True
Worksheets("RECAPITUL").Visible = False
ElseIf Worksheets("TEST").Range("R7").Value = 0 Then
Worksheets("DATACH2").Visible = True
Worksheets("RECAPITUL").Visible = False
ElseIf Worksheets("TEST").Range("R7").Value = 2 Then
Worksheets("PREDEF").Visible = True
Worksheets("RECAPITUL").Visible = False
Load UserCastype
Application.ScreenUpdating = True
UserCastype.Show
End If
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub RECAPITUL_SUIV()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("TRANSIT").Visible = True
Worksheets("RECAPITUL").Visible = False
Application.ScreenUpdating = True
Load Usercontraint
Usercontraint.Show
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub PREDIM_PREC()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("TRANSIT").Visible = True
Worksheets("PREDIM").Visible = False
Application.ScreenUpdating = True
Load Usercontraint
Usercontraint.Show
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub EQPR_PREC()
ActiveWorkbook.Unprotect Password:="max"
Worksheets("PREDIM").Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("PREDIM").Visible = True
Worksheets("EQPR").Visible = False
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub UserPredefini_PREC()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
Worksheets("ACCUEIL").Visible = True
Unload UserCastype
With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
Worksheets("PREDEF").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

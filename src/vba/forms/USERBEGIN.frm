VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Userbegin 
   Caption         =   "E@sy_PV_Projet"
   ClientHeight    =   4425
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   7150
   OleObjectBlob   =   "USERBEGIN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USERBEGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False
'On s'assure que l'utilisateur remplit correctement le userform
If OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False Or TextBox1.Value = "" Then
MsgBox "SVP, Vous devez donner un nom au projet" & Chr(10) & " mais aussi choisir un mode" & Chr(10) & _
"d 'évaluation de la consommation journalière", vbOKOnly + vbInformation, "Informations Incomplètes"
Else
With ActiveWindow
.DisplayHorizontalScrollBar = True ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = True ' SUPPRIME : Ascenseur vertical
End With
'fermeture du userform
Unload USERBEGIN
Application.StatusBar = False
Application.StatusBar = False
  If Worksheets("TEST").Range("R7") = 1 Then
'Seule la feuille de mode guidé reste visible
     
     Worksheets("DATACH1").Visible = True
     Worksheets("DATACH1").Unprotect Password:="max"
     Worksheets("DATACH1").Range("A2") = UCase("projet " & Worksheets("TEST").Range("R6"))
     Worksheets("DATACH2").Visible = False
     Worksheets("ACCUEIL").Visible = False
     efface_DATACH1
     Worksheets("DATACH1").protect Password:="max"
'     Worksheets("DATACH1").EnableSelection = xlNoSelection
     Worksheets("DATACH1").Select
    
  ElseIf Worksheets("TEST").Range("R7") = 0 Then
'Seule la feuille de mode libre reste visible
 
     Worksheets("DATACH2").Visible = True
     Worksheets("DATACH2").Unprotect Password:="max"
     Worksheets("DATACH2").Range("B2") = UCase("projet " & Worksheets("TEST").Range("R6"))
     Worksheets("DATACH1").Visible = False
     Worksheets("ACCUEIL").Visible = False
     efface_DATACH2
     Worksheets("DATACH2").protect Password:="max"
'     Worksheets("DATACH2").EnableSelection = xlNoSelection
     Worksheets("DATACH2").Select
     
    ElseIf Worksheets("TEST").Range("R7") = 2 Then
'Seule la feuille PREDEF doit être visible
Worksheets("PREDEF").Visible = True
Worksheets("ACCUEIL").Visible = False
Worksheets("PREDEF").Select
Load UserCastype
Application.ScreenUpdating = True
UserCastype.Show
   End If
End If
 Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub OptionButton1_Click()
'On indique que l'utilisateur à choisi le mode guidé
Worksheets("TEST").Range("R7") = 1
End Sub

Private Sub OptionButton2_Click()
''On indique que l'utilisateur à choisi le mode libre
Worksheets("TEST").Range("R7") = 0
End Sub

Private Sub OptionButton3_Click()
''On indique que l'utilisateur à choisi le mode Cas types
Worksheets("TEST").Range("R7") = 2
End Sub

Private Sub TextBox1_Change()
'Enregistrement direct du nom de Projet
Worksheets("TEST").Range("R6").Value = TextBox1.Value
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Usercontraint 
   Caption         =   "DONNEES DE CALCUL"
   ClientHeight    =   8055
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   8280.001
   OleObjectBlob   =   "Usercontraint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Usercontraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
Worksheets("TEST").Range("J17").Value = 1
Else
Worksheets("TEST").Range("J17").Value = 0
End If
End Sub

Private Sub ComboBox1_Change()
Worksheets("TEST").Range("J25").Value = ComboBox1.Value
End Sub

Private Sub ComboBox10_Change()
Worksheets("TEST").Range("J36").Value = ComboBox10.Value
End Sub

Private Sub ComboBox11_Change()
Worksheets("TEST").Range("J35").Value = ComboBox11.Value
End Sub

Private Sub ComboBox12_Change()
Worksheets("TEST").Range("J34").Value = ComboBox12.Value
End Sub

Private Sub ComboBox13_Change()
Worksheets("TEST").Range("J20").Value = ComboBox13.Value
End Sub

Private Sub ComboBox14_Change()
Worksheets("TEST").Range("J18").Value = ComboBox14.Value
End Sub

Private Sub ComboBox15_Change()
Worksheets("TEST").Range("J21").Value = ComboBox15.Value
End Sub

Private Sub ComboBox2_Change()
Worksheets("TEST").Range("J26").Value = ComboBox2.Value
End Sub

Private Sub ComboBox3_Change()
Worksheets("TEST").Range("J27").Value = ComboBox3.Value
End Sub

Private Sub ComboBox4_Change()
Worksheets("TEST").Range("J28").Value = ComboBox4.Value
End Sub

Private Sub ComboBox5_Change()
Worksheets("TEST").Range("J29").Value = ComboBox5.Value
End Sub

Private Sub ComboBox6_Change()
Worksheets("TEST").Range("J30").Value = ComboBox6.Value
End Sub

Private Sub ComboBox7_Change()
Worksheets("TEST").Range("J33").Value = ComboBox7.Value
End Sub

Private Sub ComboBox8_Change()
Worksheets("TEST").Range("J32").Value = ComboBox8.Value
End Sub

Private Sub ComboBox9_Change()
Worksheets("TEST").Range("J31").Value = ComboBox9.Value
End Sub

Private Sub CommandButton1_Click()
a = Worksheets("TEST").Range("M17").Value
b = Worksheets("TEST").Range("J17").Value
C = Worksheets("TEST").Range("J15").Value
If a = 0 And b = 0 Then
'Cas utilisateur ne sélectionne rien
MsgBox "Vous devez obligatoirement indiquer" & Chr(10) & "Soit la station de mesure, soit l'irradiation dans le cas SCMC et cocher la case associée" & Chr(10) & _
"Ces données sont d'une importance capitale dans le calcul du générateur PV.", vbOKOnly + vbExclamation, "Données incorrectes"
ElseIf b = 1 And Val(C) = 0 Or Val(C) < 0 Then
'Cas case cochée mais irradiation nulle
MsgBox "l'irradiation à considérer est nulle" & Chr(10) & "Veuillez entrez une irradiation valide SVP" & Chr(10) & _
"Ceci est d'une importance capital pour le calcul du Générateur PV.", vbOKOnly + vbExclamation, "Irradiation invalide"
Else
ActiveWorkbook.Unprotect Password:="max"
Usercontraint.Hide
'Calcul du générateur
PREDIMENSIONNEMENT
'Bilan production
Bilan
'Equipements proposés
PROPOSITION
Application.ScreenUpdating = False
Worksheets("PREDIM").Visible = True
Worksheets("TRANSIT").Visible = False
Worksheets("PREDIM").Select
Worksheets("PREDIM").Unprotect Password:="max"
Application.ScreenUpdating = True
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub

Private Sub CommandButton2_Click()
ActiveWorkbook.Unprotect Password:="max"
Usercontraint.Hide
Application.ScreenUpdating = False
Worksheets("RECAPITUL").Visible = True
Worksheets("TRANSIT").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton3_Click()
ActiveWorkbook.Unprotect Password:="max"
 Usercontraint.Hide
Application.ScreenUpdating = False
Worksheets("RECAPITUL").Visible = True
Worksheets("TRANSIT").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton4_Click()
a = Worksheets("TEST").Range("M17").Value
b = Worksheets("TEST").Range("J17").Value
C = Worksheets("TEST").Range("J15").Value
If a = 0 And b = 0 Then
'Cas utilisateur ne sélectionne rien
MsgBox "Vous devez obligatoirement indiquer" & Chr(10) & "Soit la station de mesure, soit l'irradiation dans le cas SCMC et cocher la case associée" & Chr(10) & _
"Ces données sont d'une importance capitale dans le calcul du générateur PV.", vbOKOnly + vbExclamation, "Données incorrectes"
ElseIf b = 1 And Val(C) = 0 Or Val(C) < 0 Then
'Cas case cochée mais irradiation nulle
MsgBox "l'irradiation à considérer est nulle" & Chr(10) & "Veuillez entrez une irradiation valide SVP" & Chr(10) & _
"Ceci est d'une importance capital pour le calcul du Générateur PV.", vbOKOnly + vbExclamation, "Irradiation invalide"
Else
ActiveWorkbook.Unprotect Password:="max"
Usercontraint.Hide
'Calcul du générateur
PREDIMENSIONNEMENT
'Bilan production
Bilan
'Equipements proposés
PROPOSITION
Application.ScreenUpdating = False
Worksheets("PREDIM").Visible = True
Worksheets("TRANSIT").Visible = False
Worksheets("PREDIM").Select
Worksheets("PREDIM").Unprotect Password:="max"
Application.ScreenUpdating = True
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub

Private Sub CommandButton5_Click()
MsgBox "Lors de la mise en place de vos modules photovoltaïques, vous devez vous assurer" _
& Chr(10) & "que ceux-ci ne risquent pas d'être masqués ou cachés par des obstacles telsque " _
& Chr(10) & " le relief,les arbres où des batiments adjacents à votre site.," & Chr(10) & _
"Les masques peuvent diminuer de moitié la production de votre générateur photovoltaïque" _
& Chr(10) & " et même l'annulée complètement." _
 & Chr(10) & "Nous vous recommandons vivement d'en prendre compte.", vbOKOnly + vbInformation, "Note sur les masques"
End Sub

Private Sub CommandButton6_Click()
MsgBox "La présente carte est celle des diverses stations solarimétriques du Cameroun." _
& Chr(10) & "Si votre site est très proche de l'une d'entre elles,ou du moins présente " _
& Chr(10) & "approximativement les mêmes caractéristiques climatiques que cette dernière," & Chr(10) & _
" alors,nous vous recommandons de la choisir comme station de référence pour le calcul de" _
& Chr(10) & " votre générateur photovoltaïque." _
& Chr(10) & "Vous pourrez alors avoir des informations sur votre production annuelle et sur" _
 & Chr(10) & "l'évolution statistique de l'irradiation de votre site sur l'année.", vbOKOnly + vbInformation, "Stations solarimétriques du Cameroun"
End Sub

Private Sub CommandButton7_Click()
a = Worksheets("TEST").Range("M17").Value
b = Worksheets("TEST").Range("J17").Value
C = Worksheets("TEST").Range("J15").Value
If a = 0 And b = 0 Then
'Cas utilisateur ne sélectionne rien
MsgBox "Vous devez obligatoirement indiquer" & Chr(10) & "Soit la station de mesure, soit l'irradiation dans le cas SCMC et cocher la case associée" & Chr(10) & _
"Ces données sont d'une importance capitale dans le calcul du générateur PV.", vbOKOnly + vbExclamation, "Données incorrectes"
ElseIf b = 1 And Val(C) = 0 Or Val(C) < 0 Then
'Cas case cochée mais irradiation nulle
MsgBox "l'irradiation à considérer est incorrecte" & Chr(10) & "Veuillez entrez une irradiation valide SVP" & Chr(10) & _
"Ceci est d'une importance capital pour le calcul du Générateur PV.", vbOKOnly + vbExclamation, "Irradiation invalide"
Else
ActiveWorkbook.Unprotect Password:="max"
Usercontraint.Hide
'Calcul du générateur
PREDIMENSIONNEMENT
'Bilan production
Bilan
'Equipements proposés
PROPOSITION
Application.ScreenUpdating = False
Worksheets("PREDIM").Visible = True
Worksheets("TRANSIT").Visible = False
Worksheets("PREDIM").Select
Worksheets("PREDIM").Unprotect Password:="max"
Application.ScreenUpdating = True
Worksheets("PREDIM").Range("A1").Select
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End If
End Sub

Private Sub CommandButton8_Click()
ActiveWorkbook.Unprotect Password:="max"
 Usercontraint.Hide
Application.ScreenUpdating = False
Worksheets("RECAPITUL").Visible = True
Worksheets("TRANSIT").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

Private Sub CommandButton9_Click()
MsgBox "         E@sy_PV ne fait pas de calcul sur la mise à la terre de votre installation." _
& Chr(10) & "Toutefois, il est important de noter que la mise à la terre des installations " _
& Chr(10) & "photovoltaïques est nécessaire pour évacuer entre autre les phénomènes transitoires," & Chr(10) & _
"dus à la foudre, car en effet, ils représentent un risque de détérioration des appareils électroniques." _
& Chr(10) & "La mise à la terre assure aussi l'équipotentialité des  masses des récepteurs (Cf Norme NFC15100)." _
& Chr(10) & "Donc, nous vous recommandons, à défaut d'une solution Paratonnerre-Parafoudre, de veiller " _
 & Chr(10) & " à ce qu'elle soit faite par les techniciens qui effecturons la mise en place de " & Chr(10) & _
 "votre générateur photovoltaïque.", vbOKOnly + vbInformation, "Note sur la mise à la terre"
End Sub

Private Sub ListBox1_Click()
If ListBox1.Value = "CAMEROUN" Then
ListBox2.RowSource = "TEST!A17:A26"
Worksheets("TEST").Range("J13").Value = ListBox1.Value
Else
a = ListBox1.Value
ListBox2.RowSource = ""
Worksheets("TEST").Range("M17").Value = 0
MsgBox "Cette unité n'est pas encore fonctionnelle" & Chr(10) & " Aucune donnée disponible à partir de la base de données" & Chr(10) & _
"Besoin de mise à jour de la base de données requis.", vbOKOnly + vbInformation, a & ":" & "Unité non fonctionnelle"
End If
End Sub

Private Sub ListBox2_Click()
Worksheets("TEST").Range("J14").Value = ListBox2.Value
If Worksheets("TEST").Range("J14").Value = 0 Then
Worksheets("TEST").Range("M17").Value = 0
Else
Worksheets("TEST").Range("M17").Value = 1
End If
End Sub

Private Sub OptionButton1_Click()
If OptionButton1.Value = True Then
Worksheets("TEST").Range("L13").Value = 0
End If
End Sub

Private Sub OptionButton2_Click()
If OptionButton2.Value = True Then
Worksheets("TEST").Range("L13").Value = 1
End If
End Sub

Private Sub OptionButton3_Click()
If OptionButton3.Value = True Then
Worksheets("TEST").Range("L13").Value = 2
End If
End Sub

Private Sub UserForm_Initialize()
Label2.Caption = UCase("projet " & Worksheets("TEST").Range("R6"))
Label10.Caption = UCase("projet " & Worksheets("TEST").Range("R6"))
ComboBox1.Value = 31
ComboBox2.Value = 28
ComboBox3.Value = 31
ComboBox4.Value = 30
ComboBox5.Value = 31
ComboBox6.Value = 30
ComboBox7.Value = 30
ComboBox8.Value = 31
ComboBox9.Value = 31
ComboBox10.Value = 31
ComboBox11.Value = 30
ComboBox12.Value = 31
ComboBox14.Value = 1
ComboBox13.Value = 50
ComboBox15.Value = 100
OptionButton2.Value = True
Worksheets("TEST").Range("J14").Value = 0
Worksheets("TEST").Range("M17").Value = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then 'Si on clique sur la croix
        Cancel = 1 'Invalide la fermeture
End If
End Sub

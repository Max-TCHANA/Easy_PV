Attribute VB_Name = "CONSOMMATION"
'Ce tableau nous permettra de faire le récapitulatifs des équipements choisis
Dim RECAPIT(1 To 100, 1 To 6) As Variant
Dim kas As Integer

Sub Consommation_guide()
Set F = Worksheets("DATACH1")
Set F1 = Worksheets("TEST")
Dim i As Integer
Dim P As Variant
Dim E As Variant
'Initialisation de la puissaance installée et de la consommation journalière
F1.Range("E2").Value = ""
F1.Range("E3").Value = ""
F1.Range("E4").Value = ""
P = 0
E = 0
kas = 1
'initialisation de recapit
For i = 1 To 100
    For j = 1 To 6
        RECAPIT(i, j) = ""
    Next j
Next i
'Parcours de la partie éclairage
'Les lignes correspondantes vont de 16 à 44 et les colonnes sont 2 pour la puissance unitaire
'4 pour nbrd et 5 pour durée
For i = 16 To 44
    If Val(F.Cells(i, 4).Value) <> 0 And Val(F.Cells(i, 5).Value) <> 0 Then
       P = P + Val(F.Cells(i, 4).Value) * F.Cells(i, 2)
       E = E + Val(F.Cells(i, 4).Value) * Val(F.Cells(i, 5).Value) * F.Cells(i, 2)
'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "Eclairage"
       RECAPIT(kas, 2) = F.Cells(i, 1).Value
       RECAPIT(kas, 3) = F.Cells(i, 2).Value
       RECAPIT(kas, 4) = F.Cells(i, 4).Value
       RECAPIT(kas, 5) = F.Cells(i, 5).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 4).Value) * Val(F.Cells(i, 5).Value) * F.Cells(i, 2)
       kas = kas + 1
    End If
Next i

'Parcours de la partie réfrégération
'Les lignes correspondantes vont de 16 à 35 et les colonnes sont 8 pour la puissance unitaire
'10 pour nbrd et 11 pour durée
'on écarte les cellules qui indique le niveau de température
For i = 16 To 35
    If i <> 16 And i <> 24 And i <> 29 And Val(F.Cells(i, 10).Value) <> 0 And Val(F.Cells(i, 11).Value) <> 0 Then
       P = P + Val(F.Cells(i, 10).Value) * F.Cells(i, 8)
       E = E + Val(F.Cells(i, 10).Value) * Val(F.Cells(i, 11).Value) * F.Cells(i, 8)
       'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "Réfré/Congel"
       RECAPIT(kas, 2) = F.Cells(i, 7).Value
       RECAPIT(kas, 3) = F.Cells(i, 8).Value
       RECAPIT(kas, 4) = F.Cells(i, 10).Value
       RECAPIT(kas, 5) = F.Cells(i, 11).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 10).Value) * Val(F.Cells(i, 11).Value) * F.Cells(i, 8)
       kas = kas + 1
    End If
Next i

'Parcours de la partie électromenager
'Les lignes correspondantes vont de 16 à 36 et les colonnes sont 14 pour la puissance unitaire
'16 pour nbrd et 17 pour durée
For i = 16 To 36
    If Val(F.Cells(i, 16).Value) <> 0 And Val(F.Cells(i, 17).Value) <> 0 Then
       P = P + Val(F.Cells(i, 16).Value) * F.Cells(i, 14)
       E = E + Val(F.Cells(i, 16).Value) * Val(F.Cells(i, 17).Value) * F.Cells(i, 14)
       'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "Electro-menager"
       RECAPIT(kas, 2) = F.Cells(i, 13).Value
       RECAPIT(kas, 3) = F.Cells(i, 14).Value
       RECAPIT(kas, 4) = F.Cells(i, 16).Value
       RECAPIT(kas, 5) = F.Cells(i, 17).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 16).Value) * Val(F.Cells(i, 17).Value) * F.Cells(i, 14)
       kas = kas + 1
    End If
Next i

'Parcours de la partie Audiovisuel
'Les lignes correspondantes vont de 16 à 24 et les colonnes sont 20 pour la puissance unitaire
'22 pour nbrd et 23 pour durée
For i = 16 To 24
    If Val(F.Cells(i, 22).Value) <> 0 And Val(F.Cells(i, 23).Value) <> 0 Then
       P = P + Val(F.Cells(i, 22).Value) * F.Cells(i, 20)
       E = E + Val(F.Cells(i, 22).Value) * Val(F.Cells(i, 23).Value) * F.Cells(i, 20)
       'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "Audiovisuel"
       RECAPIT(kas, 2) = F.Cells(i, 19).Value
       RECAPIT(kas, 3) = F.Cells(i, 20).Value
       RECAPIT(kas, 4) = F.Cells(i, 22).Value
       RECAPIT(kas, 5) = F.Cells(i, 23).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 22).Value) * Val(F.Cells(i, 23).Value) * F.Cells(i, 20)
       kas = kas + 1
    End If
Next i

'Parcours de la partie Informatique
'Les lignes correspondantes vont de 16 à 26 et les colonnes sont 26 pour la puissance unitaire
'28 pour nbrd et 29 pour durée
For i = 16 To 26
    If Val(F.Cells(i, 28).Value) <> 0 And Val(F.Cells(i, 29).Value) <> 0 Then
       P = P + Val(F.Cells(i, 28).Value) * F.Cells(i, 26)
       E = E + Val(F.Cells(i, 28).Value) * Val(F.Cells(i, 29).Value) * F.Cells(i, 26)
       'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "Informatique"
       RECAPIT(kas, 2) = F.Cells(i, 25).Value
       RECAPIT(kas, 3) = F.Cells(i, 26).Value
       RECAPIT(kas, 4) = F.Cells(i, 28).Value
       RECAPIT(kas, 5) = F.Cells(i, 29).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 28).Value) * Val(F.Cells(i, 29).Value) * F.Cells(i, 26)
       kas = kas + 1
    End If
Next i

'Nombre de ligne du tableau Récapit
kas = kas - 1
'Postage des valeurs calculée
F1.Range("E2").Value = P
F1.Range("E3").Value = E

'On s'assure que ni P ni E n'est nul
If P = 0 Or E = 0 Then
F1.Range("E4").Value = 0
Else
F1.Range("E4").Value = 1
End If
End Sub
Sub consommation_libre()
Set F = Worksheets("DATACH2")
Set F1 = Worksheets("TEST")
Dim i As Integer
Dim P As Variant
Dim E As Variant
'Initialisation de la puissaance installée et de la consommation journalière
F1.Range("E2").Value = ""
F1.Range("E3").Value = ""
F1.Range("E4").Value = ""
P = 0
E = 0
kas = 1
'initialisation de recapit
For i = 1 To 100
    For j = 1 To 6
        RECAPIT(i, j) = "__"
    Next j
Next i
'La ligne de départ est la 9, les colonnes 3 pour équipement, 4 pour puissance, 5 pour nombre
'6 pour durée
For i = 1 To 100
    If Val(F.Cells(i, 5).Value) <> 0 And Val(F.Cells(i, 6).Value) <> 0 And Val(F.Cells(i, 4).Value) <> 0 Then
       P = P + Val(F.Cells(i, 5).Value) * Val(F.Cells(i, 4))
       E = E + Val(F.Cells(i, 5).Value) * Val(F.Cells(i, 6).Value) * Val(F.Cells(i, 4))
     'Chargement du tableau pour les récapitulatifs
       RECAPIT(kas, 1) = "__"
       RECAPIT(kas, 2) = F.Cells(i, 3).Value
       RECAPIT(kas, 3) = F.Cells(i, 4).Value
       RECAPIT(kas, 4) = F.Cells(i, 5).Value
       RECAPIT(kas, 5) = F.Cells(i, 6).Value
       RECAPIT(kas, 6) = Val(F.Cells(i, 5).Value) * Val(F.Cells(i, 6).Value) * Val(F.Cells(i, 4))
       kas = kas + 1
    End If
Next i

'Nombre de ligne du tableau Récapit
kas = kas - 1
'Postage des valeurs calculée
F1.Range("E2").Value = P
F1.Range("E3").Value = E

'On s'assure que ni P ni E n'est nul
If P = 0 Or E = 0 Then
F1.Range("E4").Value = 0
Else
F1.Range("E4").Value = 1
End If
End Sub

Sub RECAPITULE()
Dim i As Integer
Dim j As Integer
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
j = 17
For i = 1 To kas
F.Cells(j, 2).Value = RECAPIT(i, 1)
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = RECAPIT(i, 2)
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = RECAPIT(i, 3)
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = RECAPIT(i, 4)
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = RECAPIT(i, 5)
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = RECAPIT(i, 6)
F.Cells(j, 7).Select
format
j = j + 1
Next i
End Sub
Sub format()
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Times New Roman"
'        .FontStyle = "Gras"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        '.ThemeColor = xlThemeColorLight1
        '.TintAndShade = 0
        '.ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
       ' .ThemeColor = xlThemeColorLight2
        '.TintAndShade = 0.599963377788629
        '.PatternTintAndShade = 0
    End With
End Sub

Sub MPTH0()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 64 To 65
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H65").Value
F1.Range("E3").Value = F1.Range("G65").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub

Sub MPTH1()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 68 To 70
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H70").Value
F1.Range("E3").Value = F1.Range("G70").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub

Sub MPTH2()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 73 To 76
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H76").Value
F1.Range("E3").Value = F1.Range("G76").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH3()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 79 To 82
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H82").Value
F1.Range("E3").Value = F1.Range("G82").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH4()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 85 To 88
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H88").Value
F1.Range("E3").Value = F1.Range("G88").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH5()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 91 To 96
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H96").Value
F1.Range("E3").Value = F1.Range("G96").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH6()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 99 To 104
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H104").Value
F1.Range("E3").Value = F1.Range("G104").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH7()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 107 To 113
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H113").Value
F1.Range("E3").Value = F1.Range("G113").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH8()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 116 To 123
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H123").Value
F1.Range("E3").Value = F1.Range("G123").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH9()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 126 To 134
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H134").Value
F1.Range("E3").Value = F1.Range("G134").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub
Sub MPTH10()
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("TEST")
Worksheets("RECAPITUL").Range("B7") = UCase(Worksheets("TEST").Range("R6"))
'le remplissage commence à la ligne 17, les colonnes concernnées sont 2,3,4,5,6 et 7
efface_RECAPITUL
F.Select
j = 17
For i = 137 To 146
F.Cells(j, 2).Value = F1.Cells(i, 1).Value
F.Cells(j, 2).Select
format
F.Cells(j, 3).Value = F1.Cells(i, 2).Value
F.Cells(j, 3).Select
format
F.Cells(j, 4).Value = F1.Cells(i, 3).Value
F.Cells(j, 4).Select
format
F.Cells(j, 5).Value = F1.Cells(i, 4).Value
F.Cells(j, 5).Select
format
F.Cells(j, 6).Value = F1.Cells(i, 5).Value
F.Cells(j, 6).Select
format
F.Cells(j, 7).Value = F1.Cells(i, 3).Value * F1.Cells(i, 4).Value * F1.Cells(i, 5).Value
F.Cells(j, 7).Select
format
j = j + 1
Next i
'Postage des valeurs calculées
F1.Range("E2").Value = F1.Range("H146").Value
F1.Range("E3").Value = F1.Range("G146").Value
'Indication de la puissance et de la consommation
F.Range("C11").Value = F1.Range("E2").Value
F.Range("C12").Value = F1.Range("E3").Value
End Sub

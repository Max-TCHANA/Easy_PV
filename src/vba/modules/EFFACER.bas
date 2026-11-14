Attribute VB_Name = "EFFACER"
Sub efface_DATACH1()
Set F = Worksheets("DATACH1")
Dim i As Integer
'Parcours de la partie éclairage
'Les lignes correspondantes vont de 16 à 44 et les colonnes sont 2 pour la puissance unitaire
'4 pour nbrd et 5 pour durée
For i = 16 To 44
    F.Cells(i, 4).Value = ""
    F.Cells(i, 5).Value = ""
Next i

'Parcours de la partie réfrégération
'Les lignes correspondantes vont de 16 à 35 et les colonnes sont 8 pour la puissance unitaire
'10 pour nbrd et 11 pour durée
'on écarte les cellules qui indique le niveau de température
For i = 16 To 35
    If i <> 16 And i <> 24 And i <> 29 Then
    F.Cells(i, 10).Value = ""
    F.Cells(i, 11).Value = ""
    End If
Next i

'Parcours de la partie électromenager
'Les lignes correspondantes vont de 16 à 36 et les colonnes sont 14 pour la puissance unitaire
'16 pour nbrd et 17 pour durée
For i = 16 To 36
    F.Cells(i, 16).Value = ""
    F.Cells(i, 17).Value = ""
Next i

'Parcours de la partie Audiovisuel
'Les lignes correspondantes vont de 16 à 24 et les colonnes sont 20 pour la puissance unitaire
'22 pour nbrd et 23 pour durée
For i = 16 To 24
    F.Cells(i, 22).Value = ""
    F.Cells(i, 23).Value = ""
Next i

'Parcours de la partie Informatique
'Les lignes correspondantes vont de 16 à 26 et les colonnes sont 26 pour la puissance unitaire
'28 pour nbrd et 29 pour durée
For i = 16 To 26
    F.Cells(i, 28).Value = ""
    F.Cells(i, 29).Value = ""
Next i
End Sub
Sub efface_DATACH2()
Set F = Worksheets("DATACH2")
F.Range("C9:F100").Value = ""
End Sub
Sub efface_RECAPITUL()
Set F = Worksheets("RECAPITUL")
F.Range("B17:G150").Clear
End Sub

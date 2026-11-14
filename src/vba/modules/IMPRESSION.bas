Attribute VB_Name = "IMPRESSION"
Sub imprim_RECAPIT()
On Error GoTo Surerreur
ActiveWorkbook.Unprotect Password:="max"
Set F = Worksheets("RECAPITUL")
Set F1 = Worksheets("IMP1")
Dim j As Integer
Application.ScreenUpdating = False
F1.Range("B2").Value = Date
F1.Range("B4").Value = "PROJET " & F.Range("B7").Value
F1.Range("D6").Value = F.Range("C11").Value
F1.Range("D7").Value = F.Range("C12").Value
F1.Range("B13:G150").Clear
'le remplissage commence à la ligne 13, les colonnes concernnées sont 2,3,4,5,6 et 7
j = 17
k = 13
Worksheets("IMP1").Visible = True
F1.Select
While F.Cells(j, 2) <> 0
F1.Cells(k, 2).Value = F.Cells(j, 2).Value
F1.Cells(k, 2).Select
format
F1.Cells(k, 3).Value = F.Cells(j, 3).Value
F1.Cells(k, 3).Select
format
F1.Cells(k, 4).Value = F.Cells(j, 4).Value
F1.Cells(k, 4).Select
format
F1.Cells(k, 5).Value = F.Cells(j, 5).Value
F1.Cells(k, 5).Select
format
F1.Cells(k, 6).Value = F.Cells(j, 6).Value
F1.Cells(k, 6).Select
format
F1.Cells(k, 7).Value = F.Cells(j, 7).Value
F1.Cells(k, 7).Select
format
j = j + 1
k = k + 1
Wend
'Lancement de l'impression une seule ligne de code !!!!!!!!!!!!!
' ExecuteExcel4Macro "PRINT(1,,,1,,,,,,,,2,,,TRUE,,FALSE)"
Worksheets("IMP1").PrintOut
'reselection de notre feuille
F.Select
Worksheets("IMP1").Visible = False
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
Application.ScreenUpdating = True

Exit Sub
Surerreur:
MsgBox "Un problème est survenu lors de l'impression." & Chr(10) & "Vous devez à partir du panneau de configuration de Window sélectionner une imprimante par défaut" & Chr(10) & _
"Nous nous excusons pour le désagrément encouru.", vbOKOnly + vbExclamation, "Aucune imprimante par défaut"
F.Select
Worksheets("IMP1").Visible = False
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
Application.ScreenUpdating = True
End Sub

Sub imprim_ETATPRO()
On Error GoTo Surerreur
ActiveWorkbook.Unprotect Password:="max"



Worksheets("IMP2").Range("A56:B68").Value = Worksheets("ETATPRO").Range("A36:B48").Value
Worksheets("IMP2").Range("A2").Value = Date
Worksheets("IMP2").Range("B4").Value = "PROJET " & UCase(Worksheets("TEST").Range("R6").Value)
Worksheets("IMP2").Range("D6").Value = Worksheets("TEST").Range("E2").Value
Worksheets("IMP2").Range("D7").Value = Worksheets("TEST").Range("E3").Value


Dim Mois(1 To 12) As Variant
Dim Irad(1 To 12) As Variant
Dim Prod(1 To 12) As Variant
'Récupération du choix
choix = Worksheets("TEST").Range("J17").Value
'Récupération du nombre de jour d'utilisation par mois
k = 1
For i = 25 To 36
    Mois(k) = Worksheets("TEST").Cells(i, 10).Value
    k = k + 1
Next i

k = 1

   'Recherche de la station choisi par l'utilisateur
    station = Worksheets("TEST").Range("J14").Value
    'recherche de l'irradiation quotidienne moyenne mensuelle pour chaque mois
    For i = 3 To 12
        If station = Worksheets("DATASOL").Cells(4, i).Value Then
           'Récupération de l'irradiation quotidienne moyenne mensuelle pour chaque mois
           For j = 5 To 16
               Irad(k) = 0.001 * Worksheets("DATASOL").Cells(j, i).Value
                
               k = k + 1
           Next j
        End If
    Next i
'.............................................................................................
'Recupération de la puissance crête effective
   Pcff = Worksheets("TEST").Range("E7").Value
    For i = 1 To 12
        Prod(i) = Irad(i) * Pcff
    Next i
'.............................................................................................
    k = 1
    Prodan = 0
    Deficitan = 0
    Pertean = 0
'Recopi des nombres de jour d'utilisation, consommation, irradiation
'<>
    For i = 26 To 37
        
        Worksheets("IMP2").Cells(i, 3).Value = Worksheets("TEST").Range("E5").Value
        Worksheets("IMP2").Cells(i, 2).Value = Irad(k)
        Worksheets("IMP2").Cells(i, 6).Value = Prod(k)
        Production = Prod(k)
'Energie A Produire EAP
        EAP = Worksheets("TEST").Range("E5").Value
        a = Production - Worksheets("TEST").Range("E5").Value
        If a <= 0 Then
'pertes
           Worksheets("IMP2").Cells(i, 4).Value = 0
'Déficit
           Worksheets("IMP2").Cells(i, 5).Value = -1 * a
        Else
'pertes
           Worksheets("IMP2").Cells(i, 4).Value = a
'Déficit
           Worksheets("IMP2").Cells(i, 5).Value = 0
        End If
        Prodan = Prodan + Mois(k) * Prod(k)
        Deficitan = Deficitan + Mois(k) * Worksheets("IMP2").Cells(i, 5).Value
        Pertean = Pertean + Mois(k) * Worksheets("IMP2").Cells(i, 4)
        k = k + 1
    Next i
    Worksheets("IMP2").Range("D45").Value = Prodan
    Worksheets("IMP2").Range("D46").Value = Deficitan
    Worksheets("IMP2").Range("D47").Value = Pertean
   'Déterminationdu mois le plus défavorable
    sent = Worksheets("IMP2").Cells(26, 5).Value
    For i = 26 To 37
        sent1 = Worksheets("IMP2").Cells(i, 5).Value
        If sent1 >= sent Then
           MoisDef = Worksheets("IMP2").Cells(i, 1).Value
           sent = sent1
        End If
    Next
    Worksheets("IMP2").Range("D49").Value = MoisDef
    
 Worksheets("IMP2").Range("F2").Value = UCase(station)

For i = 1 To 6
Worksheets("IMP2").Cells(16, i).Value = Mois(i)
Next i

For i = 1 To 6
Worksheets("IMP2").Cells(19, i).Value = Mois(6 + i)
Next i
Application.ScreenUpdating = False
Worksheets("IMP2").Visible = True
Worksheets("IMP2").Select
Worksheets("IMP2").PrintOut
Worksheets("ETATPRO").Select
Worksheets("IMP2").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

Exit Sub
Surerreur:
MsgBox "Un problème est survenu lors de l'impression." & Chr(10) & "Vous devez à partir du panneau de configuration de Window sélectionner une imprimante par défaut" & Chr(10) & _
"Nous nous excusons pour le désagrément encouru.", vbOKOnly + vbExclamation, "Aucune imprimante par défaut"
Worksheets("ETATPRO").Select
Worksheets("IMP2").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub
Sub imprim_PREDIM()
On Error GoTo Surerreur
ActiveWorkbook.Unprotect Password:="max"
For i = 8 To 13
Worksheets("IMP3").Cells(i - 2, 2).Value = Worksheets("PREDIM").Cells(i, 3).Value
Next i

For i = 19 To 29
Worksheets("IMP3").Cells(i - 2, 2).Value = Worksheets("PREDIM").Cells(i, 3).Value
Next i

For i = 35 To 37
Worksheets("IMP3").Cells(i - 2, 2).Value = Worksheets("PREDIM").Cells(i, 3).Value
Next i

For i = 43 To 51
Worksheets("IMP3").Cells(i - 3, 2).Value = Worksheets("PREDIM").Cells(i, 3).Value
Next i

For i = 57 To 60
Worksheets("IMP3").Cells(i - 3, 2).Value = Worksheets("PREDIM").Cells(i, 3).Value
Next i
Application.ScreenUpdating = False
Worksheets("IMP3").Visible = True
Worksheets("IMP3").Select
Worksheets("IMP3").PrintOut
Worksheets("PREDIM").Select
Worksheets("IMP3").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

Exit Sub
Surerreur:
MsgBox "Un problème est survenu lors de l'impression." & Chr(10) & "Vous devez à partir du panneau de configuration de Window sélectionner une imprimante par défaut" & Chr(10) & _
"Nous nous excusons pour le désagrément encouru.", vbOKOnly + vbExclamation, "Aucune imprimante par défaut"
Worksheets("PREDIM").Select
Worksheets("IMP3").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

End Sub
Sub imprim_EQPR()
On Error GoTo Surerreur
ActiveWorkbook.Unprotect Password:="max"
Worksheets("IMP4").Cells(1, 2).Value = Date
Worksheets("IMP4").Cells(3, 1).Value = "PROJET " & UCase(Worksheets("TEST").Range("R6").Value)
For i = 11 To 24
Worksheets("IMP4").Cells(i + 1, 2).Value = Worksheets("EQPR").Cells(i, 3).Value
Next i

For i = 30 To 46
Worksheets("IMP4").Cells(i + 1, 2).Value = Worksheets("EQPR").Cells(i, 3).Value
Next i

For i = 52 To 62
Worksheets("IMP4").Cells(i + 2, 2).Value = Worksheets("EQPR").Cells(i, 3).Value
Next i

For i = 68 To 82
Worksheets("IMP4").Cells(i + 2, 2).Value = Worksheets("EQPR").Cells(i, 3).Value
Next i


Worksheets("IMP4").Cells(89, 2).Value = Worksheets("EQPR").Cells(88, 3).Value

Application.ScreenUpdating = False
Worksheets("IMP4").Visible = True
Worksheets("IMP4").Select
Worksheets("IMP4").PrintOut
Worksheets("EQPR").Select
Worksheets("IMP4").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True

Exit Sub
Surerreur:
MsgBox "Un problème est survenu lors de l'impression." & Chr(10) & "Vous devez à partir du panneau de configuration de Window sélectionner une imprimante par défaut" & Chr(10) & _
"Nous nous excusons pour le désagrément encouru.", vbOKOnly + vbExclamation, "Aucune imprimante par défaut"
Worksheets("EQPR").Select
Worksheets("IMP4").Visible = False
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub


'While F.Cells(j, 2) <> 0
'F1.Cells(j, 2).Value = Left(F.Cells(j, 2).Value, 6) & "."
'F1.Cells(j, 2).Select
'format
'F1.Cells(j, 3).Value = Left(F.Cells(j, 3).Value, 6) & "."
'F1.Cells(j, 3).Select
'format
'F1.Cells(j, 4).Value = Left(F.Cells(j, 4).Value, 6) & "."
'F1.Cells(j, 4).Select
'format
'F1.Cells(j, 5).Value = Left(F.Cells(j, 5).Value, 6) & "."
'F1.Cells(j, 5).Select
'format
'F1.Cells(j, 6).Value = Left(F.Cells(j, 6).Value, 6) & "."
'F1.Cells(j, 6).Select
'format
'F1.Cells(j, 7).Value = Left(F.Cells(j, 7).Value, 6) & "."
'F1.Cells(j, 7).Select
'format
'j = j + 1
'Wend

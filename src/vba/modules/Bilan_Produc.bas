Attribute VB_Name = "Bilan_Produc"
Sub Bilan()
Worksheets("ETATPRO").Unprotect Password:="max"
Worksheets("PREDIM").Unprotect Password:="max"
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
'si case SCMC non cocher
If choix = 0 Then
   'Recherche de la station choisi par l'utilisateur
    station = Worksheets("TEST").Range("J14").Value
    'recherche de l'irradiation quotidienne moyenne mensuelle pour chaque mois
    For i = 3 To 12
        If station = Worksheets("DATASOL").Cells(4, i).Value Then
           'Récupération de l'irradiation quotidienne moyenne mensuelle pour chaque mois
           For j = 5 To 16
               Irad(k) = 0.001 * Worksheets("DATASOL").Cells(j, i).Value
                Worksheets("ETATPRO").Cells(k + 36, 2).Value = 1000 * Irad(k)
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
    For i = 3 To 14
        Worksheets("ETATPRO").Cells(12, i).Value = Mois(k)
        Worksheets("ETATPRO").Cells(20, i).Value = WorksheetFunction.Round(Worksheets("TEST").Range("E5").Value, 1)
        Worksheets("ETATPRO").Cells(18, i).Value = WorksheetFunction.Round(Irad(k), 1)
        Worksheets("ETATPRO").Cells(19, i).Value = WorksheetFunction.Round(Prod(k), 1)
        Production = WorksheetFunction.Round(Prod(k), 1)
'Energie A Produire EAP
        EAP = Worksheets("TEST").Range("E5").Value
        a = Production - WorksheetFunction.Round(Worksheets("TEST").Range("E5").Value, 1)
        If a <= 0 Then
'pertes
           Worksheets("ETATPRO").Cells(21, i).Value = 0
'Déficit
           Worksheets("ETATPRO").Cells(22, i).Value = -1 * a
        Else
'pertes
           Worksheets("ETATPRO").Cells(21, i).Value = a
'Déficit
           Worksheets("ETATPRO").Cells(22, i).Value = 0
        End If
        Prodan = Prodan + Mois(k) * WorksheetFunction.Round(Prod(k), 1)
        Deficitan = Deficitan + Mois(k) * Worksheets("ETATPRO").Cells(22, i).Value
        Pertean = Pertean + Mois(k) * Worksheets("ETATPRO").Cells(21, i)
        k = k + 1
    Next i
    Worksheets("ETATPRO").Range("E28").Value = Prodan
    Worksheets("ETATPRO").Range("E29").Value = Deficitan
    Worksheets("ETATPRO").Range("E30").Value = Pertean
    
   'Déterminationdu mois le plus défavorable
      sent = Worksheets("ETATPRO").Cells(18, 3).Value
    For i = 3 To 14
        sent1 = Worksheets("ETATPRO").Cells(18, i).Value
        If sent1 <= sent Then
           MoisDef = Worksheets("ETATPRO").Cells(17, i).Value
           sent = sent1
        End If
    Next
'    sent = Worksheets("ETATPRO").Cells(22, 3).Value
'    For i = 3 To 14
'        sent1 = Worksheets("ETATPRO").Cells(22, i).Value
'        If sent1 >= sent Then
'           MoisDef = Worksheets("ETATPRO").Cells(17, i).Value
'           sent = sent1
'        End If
'    Next
    Worksheets("ETATPRO").Range("E32").Value = MoisDef
    Worksheets("PREDIM").Range("C12").Value = MoisDef
    Worksheets("ETATPRO").Range("A2").Value = "PROJET " & UCase(Worksheets("TEST").Range("R6").Value)
 Worksheets("ETATPRO").Range("B5").Value = UCase(station)

Else
    'si case cocher
'    CAS SPECIAL POUR SITE AU CLIMAT MAL CONNU
'Pas de diagramme statistique
Worksheets("ETATPRO").Range("B37:B48").Value = 0
'On ne calcul pas la production
    Worksheets("ETATPRO").Range("E28").Value = "NC"
'On ne calcul pas les déficit
    Worksheets("ETATPRO").Range("E29").Value = "NC"
'On ne calcul pas les pertes
    Worksheets("ETATPRO").Range("E30").Value = "NC"
'Mois le plus défavorable inconnu
    Worksheets("ETATPRO").Range("E32").Value = "INCONNU"
    Worksheets("PREDIM").Range("C12").Value = "INCONNU"
'    Tableau de calcul des excédent et des déficit : non calculé NC
Worksheets("ETATPRO").Range("C18:N22").Value = "NC"
'Nombre de jour d'utilisation
Worksheets("ETATPRO").Range("C12:N12").Value = "NC"
'Non du projet
Worksheets("ETATPRO").Range("A2").Value = "PROJET " & UCase(Worksheets("TEST").Range("R6").Value)
'Station de référence aucune
 Worksheets("ETATPRO").Range("B5").Value = UCase("AUCUNE")
End If
Worksheets("ETATPRO").protect Password:="max"
Worksheets("ETATPRO").EnableSelection = xlNoSelection
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
End Sub





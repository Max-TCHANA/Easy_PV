Attribute VB_Name = "Choix_Equipmts"
Sub PROPOSITION()
Worksheets("EQPR").Unprotect Password:="max"
Worksheets("PREDIM").Unprotect Password:="max"
Set F = Worksheets("EQPR")
Set F1 = Worksheets("PREDIM")
Dim Trans(1 To 20) As Variant
Dim Sond As Variant
Dim a As Variant
Dim i As Integer
Dim j As Integer
Dim C As Variant
Dim Nbrpara As Variant

'**********************************************************************************************
'**********************************     MODULE PROPOSE     ***********************************

'Dans la BD les modules vont de la ligne 3 à la ligne 14, colonne 1(A) à 14(N)
'Récupétation de la puissance crête de module prise dans la partie contrainte d'utilisation
Pcu = Worksheets("TEST").Range("J20").Value
'Recherche de la ligne correspondante dans la base de données et enregistrement des données qui s'y trouvent
For i = 3 To 14
    If Pcu = Worksheets("DATAEQ2").Cells(i, 4).Value Then
       k = i
       For j = 1 To 14
           Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
       Next j
       Exit For
    End If
Next i

'Une fois la ligne détectée, remplissage de la partie module de la feuille EQPR
For j = 1 To 14
F.Cells(10 + j, 3).Value = Trans(j)
Next j
'prix en FCFA
F.Cells(24, 3).Value = (F.Cells(24, 3).Value) * 655


'Intensité maximale de charge, multipliation par le nombre de branche
F1.Range("C28").Value = (F.Range("C18").Value) * (F1.Range("C26").Value)
'surfafe approximative du champ
F1.Range("C29").Value = (F.Range("C20").Value) * (F1.Range("C22").Value)
'Envoi du courant max de charge vers la feuil de prédim
F1.Range("C36").Value = 1.5 * (F1.Range("C28").Value)

For j = 1 To 14
           Trans(j) = 0
       Next j
       
'récupération de la tension du générateur PV
Un = F1.Range("C35").Value


'**********************************************************************************************
'**********************************     REGULATEUR PROPOSE    ***********************************
'........................ Procédure du régulateur à choisir ...................................
'détermination du courant maximal d'entre le courant d'entré et le courant de sortie
'<>
C = WorksheetFunction.Max(F1.Range("C36").Value, F1.Range("C37").Value)
a = WorksheetFunction.Round(C, 0)
If a < C Then
C = a + 1
Else
C = a
End If

If Un = 12 Then
     '50 est la valeur maximale des courant dans cette gamme
     'On cherche dans la gamme des régulateurs 12 volts
     'la variable "choix" correspond à l'intensité choisit choisi dans la gamme des 12 volts
      If C <= 50 Then
         If C <= 6 Then choix = 6
         If 6 < C And C <= 8 Then choix = 8
         If 8 < C And C <= 10 Then choix = 10
         If 10 < C And C <= 12 Then choix = 12
         If 12 < C And C <= 15 Then choix = 15
         If 15 < C And C <= 20 Then choix = 20
         If 20 < C And C <= 30 Then choix = 30
         If 30 < C And C <= 50 Then choix = 50
         'On détecte la ligne dans la gamme 12 correspondante puis on copie les valeurs
         For i = 19 To 26
             If choix = Worksheets("DATAEQ2").Cells(i, 9).Value Then
                k = i
                For j = 1 To 16
                    Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
                Next j
                Exit For
             End If
         Next i
         
         'Une fois la ligne détectée, remplissage de la partie Régulateur de la feuille EQPR
         For j = 1 To 16
             F.Cells(29 + j, 3).Value = Trans(j)
         Next j
         'prix en FCFA
         F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
         'Nombre d'élément en parallèle recevra donc 1
         F.Cells(46, 3).Value = 1
       Else
           'Cas où Un = 12 et C>50: On adopte la solution mise en parallèle
           'On mettra en parallèle des régulateurs de 30A
           'Détermination du nombre à mettre en parallèle
           Nbrpara = C / 30
           a = WorksheetFunction.Round(Nbrpara, 0)
           If a < Nbrpara Then
              Nbrpara = a + 1
           Else
               Nbrpara = a
           End If
           'Copie des caractéristiques du régulateur de 30A
           For j = 1 To 16
               Trans(j) = Worksheets("DATAEQ2").Cells(25, j).Value
           Next j
           'ensuite remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 16
             F.Cells(29 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(46, 3).Value = Nbrpara
        End If
For j = 1 To 20
    Trans(j) = 0
Next j
   
ElseIf Un = 24 Then
        If C <= 30 Then
           If C <= 6 Then choix = 6
           If 6 < C And C <= 8 Then choix = 8
           If 8 < C And C <= 10 Then choix = 10
           If 10 < C And C <= 15 Then choix = 15
           If 15 < C And C <= 20 Then choix = 20
           If 20 < C And C <= 30 Then choix = 30
           'On détecte la ligne dans la gamme 24 correspondante puis on copie les valeurs
           For i = 28 To 33
               If choix = Worksheets("DATAEQ2").Cells(i, 9).Value Then
                  k = i
                  For j = 1 To 16
                      Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
                  Next j
                  Exit For
             End If
           Next i
         
           'Une fois la ligne détectée, remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 16
               F.Cells(29 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc 1
           F.Cells(46, 3).Value = 1
       Else
           'On cherche dans la gamme des régulateurs 24 volts
           'On met des régulateurs de 30A en parallèle
           Nbrpara = C / 30
           a = WorksheetFunction.Round(Nbrpara, 0)
           If a < Nbrpara Then
              Nbrpara = a + 1
           Else
              Nbrpara = a
           End If
           'Copie des caractéristiques du régulateur de 30A
           For j = 1 To 16
               Trans(j) = Worksheets("DATAEQ2").Cells(33, j).Value
           Next j
           'ensuite remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 16
               F.Cells(29 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(46, 3).Value = Nbrpara
      End If
           For j = 1 To 20
           Trans(j) = 0
           Next j
      
ElseIf Un = 48 Then
       'On cherche dans la gamme des régulateurs 48 volts
       If 100 <= C And C <= 140 Then
          'Copie des caractéristiques du régulateur de 140A
          For j = 1 To 16
           Trans(j) = Worksheets("DATAEQ2").Cells(33, j).Value
          Next j
        'ensuite remplissage de la partie Régulateur de la feuille EQPR
       For j = 1 To 16
           F.Cells(29 + j, 3).Value = Trans(j)
       Next j
       'prix en FCFA
       F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
       'Nombre d'élément en parallèle recevra donc 1
       F.Cells(46, 3).Value = 1
       Else
            'Mise en parallèle des régulateurs 40A
            Nbrpara = C / 40
            a = WorksheetFunction.Round(Nbrpara, 0)
            If a < Nbrpara Then
               Nbrpara = a + 1
            Else
               Nbrpara = a
            End If
            'Copie des caractéristiques du régulateur de 30A
            For j = 1 To 16
                Trans(j) = Worksheets("DATAEQ2").Cells(32, j).Value
                Next j
            'ensuite remplissage de la partie Régulateur de la feuille EQPR
            For j = 1 To 16
                F.Cells(29 + j, 3).Value = Trans(j)
            Next j
           'prix en FCFA
            F.Cells(45, 3).Value = (F.Cells(45, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(46, 3).Value = Nbrpara
       End If
End If
For j = 1 To 20
    Trans(j) = 0
Next j
'**********************************************************************************************
'**********************************     ONDULEUR PROPOSE    ***********************************
'Récupération de la puissance conseillée
'<>
Nbrpara = 0
Sond = F1.Range("C60").Value
a = WorksheetFunction.Round(Sond, 0)
If a < Sond Then
Sond = a + 1
Else
Sond = a
End If
'récupération de la tension du générateur PV
If Un = 12 Then
     '3000 est la valeur maximale des puissance dans cette gamme
     'On cherche dans la gamme des onduleurs 12 volts
     'la variable "choix" correspond à la puissance choisit dans la gamme des 12 volts
      If Pond <= 3000 Then
         If Pond <= 150 Then choix = 150
         If 150 < Sond And Sond <= 200 Then choix = 200
         If 200 < Sond And Sond <= 300 Then choix = 300
         If 300 < Sond And Sond <= 550 Then choix = 550
         If 550 < Sond And Sond <= 700 Then choix = 700
         If 700 < Sond And Sond <= 1000 Then choix = 1000
         If 1000 < Sond And Sond <= 1500 Then choix = 1500
         If 1500 < Sond And Sond <= 2000 Then choix = 2000
         If 2000 < Sond And Sond <= 3000 Then choix = 3000
         'On détecte la ligne dans la gamme 12 correspondante puis on copie les valeurs
         For i = 41 To 49
             If choix = Worksheets("DATAEQ2").Cells(i, 3).Value Then
                k = i
                For j = 1 To 14
                    Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
                Next j
                Exit For
             End If
         Next i
         'Une fois la ligne détectée, remplissage de la partie Onduleur de la feuille EQPR
         For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
         Next j
         'prix en FCFA
         F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
         'Nombre d'élément en parallèle recevra donc 1
         F.Cells(82, 3).Value = 1
      Else
          'On adopte une solution parallèle des onduleurs de 1500 VA
            Nbrpara = Sond / 1500
           a = WorksheetFunction.Round(Nbrpara, 0)
           If a < Nbrpara Then
              Nbrpara = a + 1
           Else
               Nbrpara = a
           End If
           'Copie des caractéristiques de l'onduleur de 1500 VA
           For j = 1 To 14
               Trans(j) = Worksheets("DATAEQ2").Cells(47, j).Value
           Next j
           'ensuite remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(82, 3).Value = Nbrpara
        End If
        
For j = 1 To 20
    Trans(j) = 0
Next j
   
ElseIf Un = 24 Then
      '3000 est la valeur maximale des puissance dans cette gamme
      'On cherche dans la gamme des onduleurs 24 volts
      'la variable "choix" correspond à la puissance choisit dans la gamme des 24 volts
      If Sond <= 3000 Then
         If Sond <= 200 Then choix = 200
         If 200 < Sond And Sond <= 350 Then choix = 350
         If 350 < Sond And Sond <= 700 Then choix = 700
         If 700 < Sond And Sond <= 750 Then choix = 750
         If 750 < Sond And Sond <= 1000 Then choix = 1000
         If 1000 < Sond And Sond <= 1400 Then choix = 1400
         If 1400 < Sond And Sond <= 1500 Then choix = 1500
         If 1500 < Sond And Sond <= 2000 Then choix = 2000
         If 2000 < Sond And Sond <= 3000 Then choix = 3000
         'On détecte la ligne dans la gamme 12 correspondante puis on copie les valeurs
         For i = 51 To 59
             If choix = Worksheets("DATAEQ2").Cells(i, 3).Value Then
                k = i
                For j = 1 To 14
                    Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
                Next j
                Exit For
             End If
         Next i
         'Une fois la ligne détectée, remplissage de la partie Onduleur de la feuille EQPR
         For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
         Next j
         'prix en FCFA
         F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
         'Nombre d'élément en parallèle recevra donc 1
         F.Cells(82, 3).Value = 1
      Else
          'On adopte une solution parallèle des onduleurs de 1500 VA
            Nbrpara = Sond / 1500
           a = WorksheetFunction.Round(Nbrpara, 0)
           If a < Nbrpara Then
              Nbrpara = a + 1
           Else
               Nbrpara = a
           End If
           'Copie des caractéristiques de l'onduleur de 1500 VA
           For j = 1 To 14
               Trans(j) = Worksheets("DATAEQ2").Cells(57, j).Value
           Next j
           'ensuite remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(82, 3).Value = Nbrpara
        End If
        
For j = 1 To 20
    Trans(j) = 0
Next j
   
ElseIf Un = 48 Then

     '7000 est la valeur maximale des puissance dans cette gamme
     'On cherche dans la gamme des onduleurs 48 volts
     'la variable "choix" correspond à la puissance choisit dans la gamme des 48 volts
      If Sond < 7000 Then
         If Sond <= 200 Then choix = 200
         If 200 < Sond And Sond <= 400 Then choix = 400
         If 400 < Sond And Sond <= 750 Then choix = 750
         If 750 < Sond And Sond <= 1500 Then choix = 1500
         If 1500 < Sond And Sond <= 7000 Then choix = 7000
         'On détecte la ligne dans la gamme 12 correspondante puis on copie les valeurs
         For i = 61 To 65
             If choix = Worksheets("DATAEQ2").Cells(i, 3).Value Then
                k = i
                For j = 1 To 14
                    Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
                Next j
                Exit For
             End If
         Next i
         'Une fois la ligne détectée, remplissage de la partie Onduleur de la feuille EQPR
         For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
         Next j
         'prix en FCFA
         F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
         'Nombre d'élément en parallèle recevra donc 1
         F.Cells(82, 3).Value = 1
      Else
          'On adopte une solution parallèle des onduleurs de 1500 VA
            Nbrpara = Sond / 1500
           a = WorksheetFunction.Round(Nbrpara, 0)
           If a < Nbrpara Then
              Nbrpara = a + 1
           Else
               Nbrpara = a
           End If
           
           'Copie des caractéristiques de l'onduleur de 1500 VA
           For j = 1 To 14
               Trans(j) = Worksheets("DATAEQ2").Cells(64, j).Value
           Next j
           'ensuite remplissage de la partie Régulateur de la feuille EQPR
           For j = 1 To 14
             F.Cells(67 + j, 3).Value = Trans(j)
           Next j
           'prix en FCFA
           F.Cells(81, 3).Value = (F.Cells(81, 3).Value) * 655
           'Nombre d'élément en parallèle recevra donc Nbpara
           F.Cells(82, 3).Value = Nbrpara
        End If
End If

For j = 1 To 20
    Trans(j) = 0
    Next j
'************************************************************************************************
'**********************************     BATTERIES PROPOSES    ***********************************
'Récupération de la Capacité unitaire
'<>
Cu = Worksheets("TEST").Range("J21").Value
'Recherche de la ligne correspondante dans la base de données et enregistrement des données qui s'y trouvent
For i = 69 To 86
    If Cu = Worksheets("DATAEQ2").Cells(i, 5).Value Then
       k = i
       For j = 1 To 11
           Trans(j) = Worksheets("DATAEQ2").Cells(i, j).Value
       Next j
       Exit For
    End If
Next i

'Une fois la ligne détectée, remplissage de la partie module de la feuille EQPR
For j = 1 To 11
F.Cells(51 + j, 3).Value = Trans(j)
Next j
'prix en FCFA
F.Cells(62, 3).Value = (F.Cells(62, 3).Value) * 655

'************************************************************************************************
'**********************************    PRIX TOTAL    ***********************************
'Nombre total de module
Ntm = F1.Range("C22").Value
'Prix unitaire module
Pum = F.Range("C24").Value


'Nombre total de régulateur
Ntr = F.Range("C46").Value
'Prix unitaire régulateur
Pur = F.Range("C45").Value


'Nombre total d'Onduleur
Nto = F.Range("C82").Value
'Prix unitaire Onduleur
Puo = F.Range("C81").Value


'Nombre total de Batteries
Ntb = F1.Range("C50").Value
'Prix unitaire Batterie
Pub = F.Range("C62").Value

Prixtotal = (Ntm * Pum) + (Ntr * Pur) + (Nto * Puo) + (Ntb * Pub)
F.Range("C88").Value = Prixtotal

Worksheets("EQPR").protect Password:="max"
Worksheets("EQPR").EnableSelection = xlNoSelection
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection

End Sub

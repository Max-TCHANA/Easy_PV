Attribute VB_Name = "CALC_GENERATEUR"
Sub PREDIMENSIONNEMENT()
Worksheets("PREDIM").Unprotect Password:="max"
Set F = Worksheets("TEST")
Dim i As Integer
'Puissance crête caculée
Dim Pc As Variant
'Puissance crête effective
Dim Pcff As Variant
'Energie produite
Dim Ep As Variant
'Energie consommée
Dim Ec As Variant
'Capacité calculé de la batterie
Dim Cbat As Variant
'Coefficient correcteur
Dim k As Variant
'Tension nominale du générateur
Dim Un As Variant
'Irradiation
Dim Ir As Variant



'********************* CALCUL DE PUISSANCE CRÊTE EFFECTIVE DU GENERATEUR   *******************
'**********************************************************************************************
'recupération de la consommation journalière
Ec = Worksheets("TEST").Range("E3").Value
'Choix du coefficient correcteur
choix = Worksheets("TEST").Range("L13").Value
If choix = 0 Then k = 0.8
If choix = 1 Then k = 0.65
If choix = 2 Then k = 1 / 3
'Calcul de l'énergie à produire
Ep = Ec / k
Worksheets("TEST").Range("E5").Value = Ep
'récupération de l'irradiation suivant le mode choisi: SCMC ou station
choix = Worksheets("TEST").Range("J17").Value
'si case SCMC non cocher
If choix = 0 Then
   'Recherche de la station choisi par l'utilisateur
    station = Worksheets("TEST").Range("J14").Value
    'recherche de l'irradiation quotidienne moyenne annuelle et latitude correspondantes
    For i = 3 To 12
        If station = Worksheets("DATASOL").Cells(4, i).Value Then
           Ir = 0.001 * Worksheets("DATASOL").Cells(17, i).Value
           'récupération de la latitude
           latitude = Worksheets("DATASOL").Cells(18, i).Value
        End If
    Next i
Else
    'si case cocher
    Ir = Worksheets("TEST").Range("J15").Value
    'récupération de la latitude
    latitude = Val(Worksheets("TEST").Range("J16").Value)
End If
'Calcul de la puissance crête
Pc = Ep / Ir
'Calcul de la puissance crête effective...................................................
'Détermination de la tension du système
'<>
'Récupération de la puissance unitaire de module choisi
Pcu = Worksheets("TEST").Range("J20").Value
If Pc < 500 Then Un = 12
If (500 <= Pc) And (Pc < 2000) Then Un = 24
If 2000 <= Pc Then Un = 48
'Détermination du nombre de modules en série
Nms = Un / 12
'Détermination du nombre de branche
Nb = Pc / (Nms * Pcu)
a = WorksheetFunction.Round(Nb, 0) - Nb
'Arrondi par défaut
If a < 0 Then
Nb = WorksheetFunction.Round(Nb, 0) + 1
Else
'arrondi par excès
Nb = WorksheetFunction.Round(Nb, 0)
End If
'Détermination du nombre total de module
Nt = Nb * Nms
'Puissance crête effective
Pcff = Pcu * Nt
Worksheets("TEST").Range("E7").Value = Pcff
'recupération de l'orientation
Orient = Val(Worksheets("TEST").Range("J19").Value)
'Calcul de l'inclinaison favorable
If latitude < 10 Then incli = 10
If (10 <= latitude) And (latitude < 30) Then incli = latitude
If (30 <= latitude) And (latitude < 40) Then incli = latitude + 10
If 40 <= latitude Then incli = latitude + 15
'********************************************************************************************

'*********************************    CALCUL DU REGULATEUR    ********************************
'**********************************************************************************************
'Calcul du courant de sortie du régulateur
'Récupération de la puissance installée
Pins = Worksheets("TEST").Range("E2").Value
Isortie = 1.5 * (Pins / Un)
a = WorksheetFunction.Round(Isortie, 0) - Isortie
If a < 0 Then
Isortie = WorksheetFunction.Round(Isortie, 0) + 1
Else
'arrondi par excès
Isortie = WorksheetFunction.Round(Isortie, 0)
End If
'*********************************************************************************************
'*********************************************************************************************

'*********************************    CALCUL DES BATTERIES   ********************************
'**********************************************************************************************
'récupération du nombre de jour d'autonomie
Autonomie = Worksheets("TEST").Range("J18").Value
'récupération de la capacité unitaire
Cu = Worksheets("TEST").Range("J21").Value
'Capacité de la batterie
Cbat = WorksheetFunction.Round((Ec * Autonomie) / (0.5 * Un), 1)
'calcul du taux de cyclage
EcAh = Ec / Un
Tc = WorksheetFunction.Round(EcAh / Cbat, 2)
'Calcul du nombre de branche
Nbranche = Cbat / Cu
a = WorksheetFunction.Round(Nbranche, 0)
'<>
If a < Nbranche Then
Nbranche = a + 1
Else
Nbranche = a
End If
Cbat = Nbranche * Cu
'*********************************************************************************************
'*********************************************************************************************


'*********************************    CALCUL DE L'ONDULEUR  ********************************
'**********************************************************************************************
'Récupération de la puissance installée
Pins = Worksheets("TEST").Range("E2").Value
'Puissance conseillée
Sond = 1.5 * Pins
'*********************************************************************************************
'*********************************************************************************************



'*********************************************************************************************
'********************************* POSTAGE VERS LA FEUILLE PREDIM  ***************************
'**********************************************************************************************
'Situation..............................................................
'Date
Worksheets("PREDIM").Range("C8").Value = Date
'Nom du projet
Worksheets("PREDIM").Range("C9").Value = UCase(Worksheets("TEST").Range("R6").Value)
'Station solarimérique
choix = Worksheets("TEST").Range("J17").Value
If choix = 0 Then
Worksheets("PREDIM").Range("C10").Value = station
Else
Worksheets("PREDIM").Range("C10").Value = "AUCUNE"
End If
'Latitude du lieu
Worksheets("PREDIM").Range("C11").Value = latitude
'Mois le plus défavorable
Worksheets("PREDIM").Range("C12").Value = "En attente"
'Besoins moyen
Worksheets("PREDIM").Range("C13").Value = Ec

'Dimensionnement du champ PV...............................................................
'Puissance crête
Worksheets("PREDIM").Range("C19").Value = Pcff
'Puissance unitaire d'un module
Worksheets("PREDIM").Range("C20").Value = Pcu
'Tension unitaire d'un module
Worksheets("PREDIM").Range("C21").Value = 12
'Nombre total de modules
Worksheets("PREDIM").Range("C22").Value = Nt
'Orientation
Worksheets("PREDIM").Range("C23").Value = Orient
'Inclinaison la plus favorable
Worksheets("PREDIM").Range("C24").Value = incli
'Nombre de module en série
Worksheets("PREDIM").Range("C25").Value = Nms
'Nombre de branches parallèles
Worksheets("PREDIM").Range("C26").Value = Nb
Worksheets("PREDIM").Range("C27").Value = Un
'Dimensionnement du régulateur.........................................................
'Tension nominale
Worksheets("PREDIM").Range("C35").Value = Un
'Courant d'entrée
Worksheets("PREDIM").Range("C36").Value = "Fonction du type de module"
'Courant de sortie
Worksheets("PREDIM").Range("C37").Value = WorksheetFunction.Round((Pins / Un), 1)

'Dimensionnement des batteries...........................................................
'Nombre de jours d'autonomie
Worksheets("PREDIM").Range("C43").Value = Autonomie
'Tension de la batterie
Worksheets("PREDIM").Range("C44").Value = Un
'Capacité de la batterie
Worksheets("PREDIM").Range("C45").Value = Cbat
'Taux de cyclage nominal
Worksheets("PREDIM").Range("C46").Value = Tc
'Capacité unitaire
Worksheets("PREDIM").Range("C47").Value = Worksheets("TEST").Range("J21").Value
'Tension unitaire
Worksheets("PREDIM").Range("C48").Value = 12

'Nombre d'éléments en série
Worksheets("PREDIM").Range("C49").Value = Nms
'Nombre de branches
Worksheets("PREDIM").Range("C50").Value = Nbranche
'Nombre total d'élément
Worksheets("PREDIM").Range("C51").Value = Nms * Nbranche
Worksheets("PREDIM").Range("C52").Value = Worksheets("TEST").Range("J18").Value

'Dimensionnement de l'Onduleur............................................................
'Tension d'entrée
Worksheets("PREDIM").Range("C57").Value = Un
'Tension de sortie
Worksheets("PREDIM").Range("C58").Value = 230
'Puissance installée
Worksheets("PREDIM").Range("C59").Value = Pins
'Puissance conseillée
Worksheets("PREDIM").Range("C60").Value = 1.5 * Pins

'*********************************************************************************************
Worksheets("PREDIM").protect Password:="max"
Worksheets("PREDIM").EnableSelection = xlNoSelection
End Sub

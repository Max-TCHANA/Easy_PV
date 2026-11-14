Attribute VB_Name = "IDENTIFICATION"
  Sub IDENTITE()
'  Application.Caption = "E@sy_PV"
'  ActiveWorkbook.Windows(1).Caption = "Outil d'aide au dimensionnement des systèmes PV"
  NomClasseur = ActiveWorkbook.Name
  Set F2 = Worksheets("TEST")
  Set F = Worksheets("ACCUEIL")
  Workbooks(NomClasseur).Worksheets("ACCUEIL").Activate
  F2.Range("R1").Value = 11

  Do
  P = 1
  Load USERPASS
  USERPASS.Show
  Login = F2.Range("R2").Value
  Pass = F2.Range("R3").Value
  sent = F2.Range("R1").Value
  Login = Login & Pass & sent & "."
  
  If sent = 11 Then
     P = MsgBox("Voulez-vous quitter E@sy_PV Administration?", vbYesNo + vbQuestion, "E@sy_PV")
    If P = 6 Then
    Unload USERPASS
    Exit Sub
'       Workbooks(NomClasseur).Close SaveChanges:=False
'       Application.Windows(1).Close
    End If
  End If
  
  If (Login <> "max1001.") And (Login <> "rosa1011.") And (Login <> "william1021.") And P <> 7 Then
     MsgBox "E@sy_PV ne reconnait pas cette identité" & Chr(13) & "ESSAYEZ A NOUNEAU", vbInformation, "E@sy_PV"
  End If
  
   F2.Range("R2").Value = ""
   F2.Range("R3").Value = ""
   F2.Range("R1").Value = 11
   Loop While (Login <> "max1001.") And (Login <> "william1021.") And (Login <> "rosa1011.")
   
   F2.Range("R2").Clear
   F2.Range("R3").Clear
   F2.Range("R1").Clear
   Workbooks(NomClasseur).Worksheets("ACCUEIL").Activate
   Load UserMotdepasse
   UserMotdepasse.Show
   End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
ActiveWorkbook.Unprotect Password:="max"
Application.ScreenUpdating = False

sent = Worksheets("TEST").Range("M68").Value
If sent = 0 Then
   With ActiveWindow
   .DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
   .DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
   End With
   Worksheets("ACCUEIL").Visible = True
   Worksheets("GUIDE").Visible = False
   Worksheets("ACCUEIL").Select
   
   Unload UserAbout
ElseIf sent = 1 Then
       Worksheets("DATACH1").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("DATACH1").Select

ElseIf sent = 2 Then
       Worksheets("DATACH2").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("DATACH2").Select
       
ElseIf sent = 3 Then
       Worksheets("RECAPITUL").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("RECAPITUL").Select
       
ElseIf sent = 4 Then
       Worksheets("PREDIM").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("PREDIM").Select
       
ElseIf sent = 5 Then
       Worksheets("ETATPRO").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("ETATPRO").Select
ElseIf sent = 6 Then
       Worksheets("EQPR").Visible = True
       Worksheets("GUIDE").Visible = False
       Worksheets("EQPR").Select
End If
   Application.ScreenUpdating = True
   ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub

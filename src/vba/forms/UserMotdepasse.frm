VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserMotdepasse 
   Caption         =   "Tous les mots de passe"
   ClientHeight    =   2700
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   4690
   OleObjectBlob   =   "UserMotdepasse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserMotdepasse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload UserMotdepasse
End Sub

Private Sub ListBox1_Click()
If ListBox1.Value = "Classeur" Then
ListBox2.RowSource = "Test!E17"
ElseIf ListBox1.Value = "Feuilles" Then ListBox2.RowSource = "Test!E18"
ElseIf ListBox1.Value = "Code Source" Then ListBox2.RowSource = "Test!E19"
End If
End Sub


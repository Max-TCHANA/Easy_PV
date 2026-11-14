VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserEquipement 
   Caption         =   "Equipements"
   ClientHeight    =   8415.001
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   7950
   OleObjectBlob   =   "UserEquipement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserEquipement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox2_Click()
If ListBox2.Value = "ETSOLAR" Then
Image1.Visible = True
Image5.Visible = False
Image6.Visible = False
ElseIf ListBox2.Value = "LORENTZ" Then
Image1.Visible = False
Image5.Visible = True
Image6.Visible = False
ElseIf ListBox2.Value = "SUNTECH" Then
Image1.Visible = False
Image5.Visible = False
Image6.Visible = True
End If
End Sub

Private Sub ListBox3_Click()
If ListBox3.Value = "SBC" Then
Image2.Visible = True
Image7.Visible = False
Image8.Visible = False
Image13.Visible = False
ElseIf ListBox3.Value = "SMART" Then
Image2.Visible = False
Image7.Visible = True
Image8.Visible = False
Image13.Visible = False
ElseIf ListBox3.Value = "STECA" Then
Image2.Visible = False
Image7.Visible = False
Image8.Visible = True
Image13.Visible = False
ElseIf ListBox3.Value = "Tarom" Then
Image2.Visible = False
Image7.Visible = False
Image8.Visible = False
Image13.Visible = True
End If
End Sub

Private Sub ListBox6_Click()
If ListBox6.Value = "ACDIS" Then
Image3.Visible = True
Image9.Visible = False
Image10.Visible = False
ElseIf ListBox6.Value = "Banner" Then
Image3.Visible = False
Image9.Visible = True
Image10.Visible = False
ElseIf ListBox6.Value = "STECO" Then
Image3.Visible = False
Image9.Visible = False
Image10.Visible = True
End If
End Sub

Private Sub ListBox8_Click()
If ListBox8.Value = "AL" Then
Image4.Visible = True
Image11.Visible = False
Image12.Visible = False
ElseIf ListBox8.Value = "HPC" Then
Image4.Visible = False
Image11.Visible = True
Image12.Visible = False
ElseIf ListBox8.Value = "Sinepower" Then
Image4.Visible = False
Image11.Visible = False
Image12.Visible = True
End If
End Sub

Private Sub UserForm_Initialize()
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False
Image8.Visible = False
Image9.Visible = False
Image10.Visible = False
Image11.Visible = False
Image12.Visible = False
Image13.Visible = False
End Sub

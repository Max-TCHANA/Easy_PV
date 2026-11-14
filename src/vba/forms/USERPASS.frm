VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USERPASS 
   Caption         =   "E@sy_PV Administration"
   ClientHeight    =   2385
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   6600
   OleObjectBlob   =   "USERPASS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USERPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_Change()
Set F2 = Worksheets("TEST")
F2.Range("R2") = TextBox1
End Sub
Private Sub UserForm_Initialize()
Label_Date.Caption = Date
End Sub

Private Sub CommandButton1_Click()
Set F2 = Worksheets("TEST")
F2.Range("R1").Value = 1
Unload USERPASS
End Sub

Private Sub TextBox2_Change()
Set F2 = Worksheets("TEST")
F2.Range("R3") = TextBox2
End Sub


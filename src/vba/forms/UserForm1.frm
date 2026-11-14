VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5565
   ClientLeft      =   50
   ClientTop       =   440
   ClientWidth     =   9960.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload UserForm1
'UserForm1.Hide

End Sub

Private Sub ListBox1_Click()
Dim i As Integer, j As Integer
    Dim Val As String
'    j = ListBox1.ListIndex  'renvoie 4
    Val = ListBox1.Value    'renvo
    
    If Val = "Janvier" Then
    UserForm1.ListBox2.RowSource = "Test!A3:A14"
    i = ListBox2.ListCount
    MsgBox i, vbOKOnly + vbInformation, Petittest
    Else
   UserForm1.ListBox2.RowSource = ""
   i = ListBox2.ListCount
   MsgBox i, vbOKOnly, vbInformation
'     "Test!E1:E2"
    End If
End Sub

Private Sub UserForm_Initialize()
Label1.Caption = "MAISON AUTONOME"
'UserForm1.ListBox1
'UserForm.AddItem "FRED", 2
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
On Error GoTo Surerreur
NomClasseur = ActiveWorkbook.Name
CacheBOutils
Userbeg
Exit Sub
Surerreur:
MsgBox "Une Erreur système c'est produite" & Chr(10) & "E@sy_PV doit se fermer pour se préserver,vous pouvz redémarrer par la suite" & Chr(10) & _
"Nous nous excusons pour le désagrément.", vbOKOnly + vbCritical, "Erreur système !!!"
Workbooks(NomClasseur).Close SaveChanges:=False
Application.Windows(1).Close
End Sub

Private Sub CommandButton2_Click()
CacheBOutils
IDENTITE
End Sub

Private Sub CommandButton3_Click()
CacheBOutils
Load UserAbout
UserAbout.Show
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
On Error GoTo Surerreur
Application.Caption = "E@sy_PV se Ferme"
ActiveWorkbook.Windows(1).Caption = "E@sy_PV se Ferme"
Application.StatusBar = False
Application.StatusBar = False
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
NomClasseur = ActiveWorkbook.Name
'save true pour prendre compte de la fermeture des feuilles actives et de l'ouverture de acceuil
Ouverture_classeur
'False
Application.Caption = "Microsoft Excel"
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
Application.DisplayFormulaBar = True
Application.DisplayStatusBar = True
Application.StatusBar = False
AfficheBOutils
Workbooks(NomClasseur).Close SaveChanges:=True
Application.Windows(1).Close
Exit Sub
Surerreur:
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
ActiveWorkbook.Unprotect Password:="max"
NomClasseur = ActiveWorkbook.Name
Application.StatusBar = False
Application.StatusBar = False
Application.StatusBar = False
MsgBox "Une erreur c'est produite." & Chr(10) & "E@sy_PV est surement utilisé à partir d'un support ROM" & Chr(10) & _
"Veuillez faire une copie sur votre machine" & Chr(10) & "Nous nous excusons pour le désagrément.", vbOKOnly + vbCritical, "Erreur système !!!"
Workbooks(NomClasseur).Close SaveChanges:=False
Application.Windows(1).Close
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
On Error GoTo Surerreur
NomClasseur = ActiveWorkbook.Name
Exit Sub
Surerreur:
MsgBox "Une Erreur système c'est produite" & Chr(10) & "E@sy_PV doit se fermer pour se préserver,vous pouvz redémarrer par la suite" & Chr(10) & _
"Nous nous excusons pour le désagrément.", vbOKOnly + vbCritical, "Erreur système !!!"
Workbooks(NomClasseur).Close SaveChanges:=False
Application.Windows(1).Close
End Sub

Private Sub Workbook_Open()
On Error GoTo Surerreur
  CacheBOutils
  With ActiveWindow
.DisplayHorizontalScrollBar = False ' SUPPRIME : Ascenseur horizontal
.DisplayVerticalScrollBar = False ' SUPPRIME : Ascenseur vertical
End With
  Application.StatusBar = "Bienvenue sur E@sy_PV"
  Application.DisplayFormulaBar = False
  Application.Caption = "E@sy_PV"
  ActiveWorkbook.Windows(1).Caption = "Outil d'aide au dimensionnement des systèmes PV"
Application.WindowState = xlMaximized
Application.WindowState = xlMaximized
Application.DisplayFormulaBar = False
RECEPT
Ouverture_classeur
Application.DisplayStatusBar = False
Exit Sub
Surerreur:
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
MsgBox "Une erreur c'est produite au démarrage." & Chr(10) & "Cette erreur survient géralement lorsque plusieurs classeurs excel sont ouvert" & Chr(10) & _
"Veuillez fermer tout les classeurs ouverts, y compris E@sy_PV" & Chr(10) & " Puis relancer Easy_PV en s'assurant qu'il est le seul classeur Excel ouvert" & Chr(10) & "Nous nous excusons pour le désagrément.", vbOKOnly + vbCritical, "Erreur système !!!"
End Sub

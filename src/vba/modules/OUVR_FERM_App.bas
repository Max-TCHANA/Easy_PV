Attribute VB_Name = "OUVR_FERM_App"
Sub Ouverture_classeur()
Application.ScreenUpdating = False
ActiveWorkbook.Unprotect Password:="max"
Worksheets("ACCUEIL").Visible = True
Worksheets("ACCUEIL").Select
For Each sh In Sheets
If sh.Name <> "ACCUEIL" Then
 sh.Visible = False
 End If
Next sh
Application.ScreenUpdating = True
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
End Sub


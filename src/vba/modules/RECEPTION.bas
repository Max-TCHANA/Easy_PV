Attribute VB_Name = "RECEPTION"
Sub RECEPT()
Application.WindowState = xlMaximized
Application.WindowState = xlMaximized
Set F1 = Worksheets("ACCUEIL")
Worksheets("ACCUEIL").Unprotect Password:="max"
F1.Select
Worksheets("ACCUEIL").Unprotect Password:="max"
On Error Resume Next
Cancel = True
ActiveWindow.DisplayGridlines = False

ActiveSheet.DrawingObjects.Select
'Selection.Delete

'ActiveSheet.Shapes.AddTextEffect(msoTextEffect17, "E@sy_PV", _
'"Times New Roman", 130#, msoFalse, msoFalse, 45, 1).Select
''"Times New Roman"
'Selection.Placement = xlFreeFloating
'un = Selection.Name
'ActiveSheet.Shapes.AddTextEffect(msoTextEffect27, "Calcul-Propositions-Rapports", "Comic Sans MS", _
'25#, msoFalse, msoFalse, 230, 170).Select
'Comic Sans MS
'Selection.Placement = xlFreeFloating
'deux = Selection.Name
ActiveSheet.Shapes(Un).Select
m = 1

For i = 1 To 5
Selection.ShapeRange.TextEffect.Tracking = m
m = m + 0.1
DoEvents
Next i

For i = 5 To 1 Step -1
Selection.ShapeRange.TextEffect.Tracking = m
m = m - 0.1
DoEvents
Next i

Selection.ShapeRange.TextEffect.Tracking = 1
For k = 1 To 2
For i = 0 To 1 Step 0.1
[A200].Value = i
Calculate

Next i
[A300].Value = 1
For j = 1 To 0 Step -0.1
[A200].Value = j
Calculate

Next j
Next k

ActiveSheet.Shapes(deux).Select

For i = 1 To 12
'Selection.ShapeRange.IncrementRotation 30
DoEvents
Next i

'ActiveSheet.Shapes.AddTextEffect(msoTextEffect17, "Outil d'aide au dimensionnement des systèmes Photovoltaïques", _
'"Arial Black", 18#, msoFalse, msoFalse, 40, 240).Select
Selection.Placement = xlFreeFloating
[A1].Select
Worksheets("ACCUEIL").protect Password:="max", DrawingObjects:=True, Contents:=True, Scenarios:=True
'"Arial Black"
End Sub



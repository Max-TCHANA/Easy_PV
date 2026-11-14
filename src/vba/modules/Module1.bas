Attribute VB_Name = "Module1"
Sub Protéger()
Attribute Protéger.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Protéger Macro
'

'
ActiveWorkbook.protect Password:="max", Structure:=True, Windows:=True
ActiveWorkbook.Unprotect Password:="max"
End Sub
'"Suppression" (je mets des guillemets !) :
Sub CacheBOutils()
Dim CmdB As CommandBar
  For Each CmdB In Application.CommandBars
    CmdB.Enabled = False
  Next CmdB
Application.DisplayFormulaBar = False
Application.CommandBars("Worksheet Menu Bar").Enabled = False
End Sub

'Résurrection:
Sub AfficheBOutils()
Dim CmdB As CommandBar
  For Each CmdB In Application.CommandBars
    CmdB.Enabled = True
  Next CmdB
End Sub

Attribute VB_Name = "ModCPainel"
Option Compare Database
Function painelcontrole()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Function

Attribute VB_Name = "ModRede"
Option Compare Database

Function mostrarede()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Function

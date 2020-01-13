Attribute VB_Name = "ModPropsistema"
Option Compare Database
Function propsistema()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Function


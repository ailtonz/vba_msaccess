VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Comando53_Click()
Call painelcontrole
End Sub

Private Sub Comando54_Click()
Call propsistema
End Sub

Private Sub Comando55_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub


Private Sub Comando57_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub Comando58_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub Comando59_Click()
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub



  

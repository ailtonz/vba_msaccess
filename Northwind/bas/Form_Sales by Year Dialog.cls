VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sales by Year Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Private Sub Cancel_Click()
' This code created by Command Button Wizard.
On Error GoTo Err_Cancel_Click

    ' Close form.
    DoCmd.Close

Exit_Cancel_Click:
    Exit Sub

Err_Cancel_Click:
    MsgBox Err.Description
    Resume Exit_Cancel_Click


End Sub





Private Sub OK_Click()
On Error GoTo Err_OK_Click

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer
    
    ' If Sales by Year report is not being opened for previewing or printing,
    ' cause an error. (blnOpening variable is true only when report's Open event
    ' is being executed.)
    If Not Reports![Sales By Year].blnOpening Then Err.Raise 0
    
    ' Hide form.
    Me.Visible = False

Exit_OK_Click:
    Exit Sub

Err_OK_Click:
    strMsg = "To use this form, you must preview or print the Sales by Year report from the Database window or Design view."
    intStyle = vbOKOnly
    strTitle = "Open from Report"

    MsgBox strMsg, intStyle, strTitle
    Resume Exit_OK_Click

End Sub



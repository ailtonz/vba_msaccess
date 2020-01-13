VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.


Private Sub OutputToHTML_Click()
On Error GoTo Err_OuputToHTML_Click
' Outputs the Alphabetical List of Products report as an HTML
' document, and opens the document in an Internet browser.
' You must have Nwindtem.htm (template for Northwind) and NWLogo.gif
' (Northwind logo) in your default database folder.
    
    DoCmd.OutputTo acOutputReport, "Alphabetical List of Products", acFormatHTML, "Products.htm", True, "Nwindtem.htm"
    
Exit_OutputToHTML_Click:
    Exit Sub
    
Err_OuputToHTML_Click:
    ' If action was cancelled by the user, don't display an error message.
    Const conErrDoCmdCancelled = 2501
    If (Err = conErrDoCmdCancelled) Then
        Resume Exit_OutputToHTML_Click
    Else
        MsgBox Err.Description
        Resume Exit_OutputToHTML_Click
    End If
    
End Sub

Private Sub ProductName_AfterUpdate()
' If OpenArgs property isn't null, set SupplierID to value of form's OpenArgs
' property. OpenArgs will have a value if Products form is opened by clicking
' AddProducts command button on Suppliers form.
    If IsNull(Forms!Products.OpenArgs) Then
        Exit Sub
    Else
        Me!SupplierID = Forms!Products.OpenArgs
    End If
    
End Sub

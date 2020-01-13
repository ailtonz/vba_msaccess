VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Orders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.


Private Sub CustomerID_AfterUpdate()
' Update ShipTo controls based on value selected in CustomerID combo box.
    Me!ShipName = Me![CustomerID].Column(1)
    Me!ShipAddress = Me!Address
    Me!ShipCity = Me!City
    Me!ShipRegion = Me!Region
    Me!ShipPostalCode = Me!PostalCode
    Me!ShipCountry = Me!Country

End Sub


Private Sub CustomerID_BeforeUpdate(Cancel As Integer)
' Display message if CustomerID combo box is blank.

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer
    
    If IsNull(Me!CustomerID) Or Me!CustomerID = "" Then
        strMsg = "You must pick a value from the Bill To list."
        strTitle = "Bill To Customer Required"
        intStyle = vbOKOnly
        MsgBox strMsg, intStyle, strTitle
        Cancel = True
    End If

End Sub



Sub PrintInvoice_Click()
' This code created by Command Button Wizard.
On Error GoTo Err_PrintInvoice_Click

    Dim strDocName As String
    
    strDocName = "Invoice"
    ' Print Invoice report, using Invoices Filter query to print
    ' invoice for current order.
    DoCmd.OpenReport strDocName, acViewNormal, "Invoices Filter"

Exit_PrintInvoice_Click:
    Exit Sub

Err_PrintInvoice_Click:
    ' If action was cancelled by the user, don't display an error message.
    Const conErrDoCmdCancelled = 2501
    If (Err = conErrDoCmdCancelled) Then
        Resume Exit_PrintInvoice_Click
    Else
        MsgBox Err.Description
        Resume Exit_PrintInvoice_Click
    End If

End Sub

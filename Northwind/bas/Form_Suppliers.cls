VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Suppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Sub AddProducts_Click()
' This code created in part by Command Button Wizard.
On Error GoTo Err_AddProducts_Click

    Dim strDocName As String

    strDocName = "Products"
    ' Open Products form in data entry mode and store SupplierID in
    ' the form's OpenArgs property.
    DoCmd.OpenForm strDocName, , , , acAdd, , Me!SupplierID
    
    ' Close Product List form.
    DoCmd.Close acForm, "Product List"

    ' Give ProductName control focus.
    Forms![Products]!ProductName.SetFocus
    
Exit_AddProducts_Click:
    Exit Sub
    
Err_AddProducts_Click:
    MsgBox Err.Description
    Resume Exit_AddProducts_Click

End Sub


Private Sub Form_BeforeUpdate(Cancel As Integer)

' If number of digits entered in PostalCode text box is incorrect for value
' in Country text box, display message and undo PostalCode value.

    Select Case Me!Country
        Case IsNull(Me![Country])
            Exit Sub
        Case "France", "Italy", "Spain"
            If Len(Me![PostalCode]) <> 5 Then
                MsgBox "Postal Code must be 5 characters", 0, "Postal Code Error"
                Cancel = True
                Me![PostalCode].SetFocus
            End If
        Case "Australia", "Singapore"
            If Len(Me![PostalCode]) <> 4 Then
                MsgBox "Postal Code must be 4 characters", 0, "Postal Code Error"
                Cancel = True
                Me![PostalCode].SetFocus
            End If
        Case "Canada"
            If Not Me![PostalCode] Like "[A-Z][0-9][A-Z] [0-9][A-Z][0-9]" Then
                MsgBox "Postal Code not valid. Example of Canadian code: H1J 1C3", 0, "Postal Code Error"
                Cancel = True
                Me![PostalCode].SetFocus
            End If
            
    End Select


End Sub

Private Sub Form_Close()

    ' Close Product List form and Products form if they are open.
    If IsLoaded("Product List") Then DoCmd.Close acForm, "Product List"
    If IsLoaded("Products") Then DoCmd.Close acForm, "Products"
    
End Sub

Private Sub Form_Current()
On Error GoTo Err_Form_Current

' If Product List form is open, show current supplier's products.

    Dim strDocName As String
    Dim strLinkCriteria As String

        strDocName = "Product List"
        strLinkCriteria = "[SupplierID] = Forms![Suppliers]![SupplierID]"
   
    If IsNull(Me![CompanyName]) Then
        Exit Sub
    ElseIf IsLoaded("Product List") Then
        DoCmd.OpenForm strDocName, , , strLinkCriteria
    End If

Exit_Form_Current:
    Exit Sub
    
Err_Form_Current:
    MsgBox Err.Description
    Resume Exit_Form_Current

End Sub



Private Sub ReviewProducts_Click()
' This code created in part by Command Button Wizard.
On Error GoTo Err_ReviewProducts_Click

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer
    Dim strDocName As String, strLinkCriteria As String

    ' If CompanyName control is blank, display a message.
    If IsNull(Me![CompanyName]) Then
        strMsg = "Move to the supplier record whose products you want to see, then press the Review Products button again."
        intStyle = vbOKOnly
        strTitle = "Select a Supplier"
        MsgBox strMsg, intStyle, strTitle
        Me![CompanyName].SetFocus
    Else
    ' Otherwise, open Product List form, showing products for current supplier.
        strDocName = "Product List"
        strLinkCriteria = "[SupplierID] = Forms![Suppliers]![SupplierID]"
        DoCmd.OpenForm strDocName, , , strLinkCriteria
        DoCmd.MoveSize (1440 * 0.78), (1440 * 1.8)
    End If
    
Exit_ReviewProducts_Click:
    Exit Sub
    
Err_ReviewProducts_Click:
    MsgBox Err.Description
    Resume Exit_ReviewProducts_Click
    
End Sub




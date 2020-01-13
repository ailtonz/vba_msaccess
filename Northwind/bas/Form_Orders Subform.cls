VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Orders Subform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
' If Orders Subform is opened as a standalone form, display a message
' and undo the deletion. (The IsLoaded function finds a form in the
' Forms collection when it's opened as a standalone form, but not when
' it's opened as a subform on a main form.)

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer

    If IsLoaded("Orders Subform") Then
        strMsg = "You can't delete a record when you open Orders Subform as a standalone form."
        intStyle = vbOKOnly
        strTitle = "Can't Delete Record"
        MsgBox strMsg, intStyle, strTitle
        Cancel = True
    End If

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
' If Orders Subform is opened as a standalone form, display a message
' and undo changes made to the ProductName text box. (The IsLoaded function
' finds a form in the Forms collection when it's opened as a standalone
' form, but not when it's opened as a subform on a main form.)

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer

    If IsLoaded("Orders Subform") Then
        strMsg = "You can't add or edit a record when you open Orders Subform as a standalone form."
        intStyle = vbOKOnly
        strTitle = "Can't Add or Change Record"
        MsgBox strMsg, intStyle, strTitle
        Me.Undo
    End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
' Errors can occur here if the user has not already selected a customer for
' this order on the main form. If there is an error and Orders!BillTo is Null
' inform the user and cancel the entry made in the subform.
    If DataErr > 0 Then
        If IsNull(Me.Parent!CustomerID) Then
            MsgBox "Select a Customer to bill to before entering order details info."
            RunCommand acCmdUndo
            Me.Parent!CustomerID.SetFocus
            Response = acDataErrContinue
        Else
            Response = acDataErrDisplay
    End If
End If

End Sub

Private Sub ProductID_AfterUpdate()
On Error GoTo Err_ProductID_AfterUpdate

    Dim strFilter As String
    
    ' Evaluate filter before it's passed to DLookup function.
    strFilter = "ProductID = " & Me!ProductID
    
    ' Look up product's unit price and assign it to UnitPrice control.
    Me!UnitPrice = DLookup("UnitPrice", "Products", strFilter)

Exit_ProductID_AfterUpdate:
    Exit Sub

Err_ProductID_AfterUpdate:
    MsgBox Err.Description
    Resume Exit_ProductID_AfterUpdate

End Sub


Private Sub ProductID_BeforeUpdate(Cancel As Integer)
' If Orders Subform is opened as a standalone form, display a message
' and undo changes made to the ProductName text box. (The IsLoaded function
' finds a form in the Forms collection when it's opened as a standalone
' form, but not when it's opened as a subform on a main form.)

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer
    
    If IsLoaded("Orders Subform") Then
        strMsg = "You can't add or edit a Product Name when you open Orders Subform as a standalone form."
        intStyle = vbOKOnly
        strTitle = "Can't Add or Change Product Name"
        MsgBox strMsg, intStyle, strTitle
        Me!ProductID.Undo
        Me.Undo
    End If

End Sub





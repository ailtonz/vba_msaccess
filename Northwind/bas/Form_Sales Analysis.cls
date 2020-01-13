VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sales Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnEdit_Click()
On Error GoTo btnEdit_Err
    If Me.Sales_Analysis_Subform1.SourceObject = "Sales Analysis Subform1" Then
        Me.Sales_Analysis_Subform1.SourceObject = "Sales Analysis Subform2"
        Me.btnEdit.Caption = "Show PivotTable"
    Else
        Me.Sales_Analysis_Subform1.SourceObject = "Sales Analysis Subform1"
        Me.btnEdit.Caption = "Show PivotChart"
    End If

btnEdit_Exit:
    Exit Sub
btnEdit_Err:
    MsgBox Err.Description
    Resume btnEdit_Exit
End Sub

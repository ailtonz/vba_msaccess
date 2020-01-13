VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sales Reports Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.


Sub PrintReports(PrintMode As Integer)
On Error GoTo Err_Preview_Click
    ' This procedure used in Preview_Click and Print_Click Sub procedures.
    ' Preview or print report selected in the ReportToPrint option group.
    ' Then close the Print Sales Reports Dialog form.

    Dim strWhereCategory As String

    strWhereCategory = "CategoryName = Forms![Sales Reports Dialog]!SelectCategory"

    Select Case Me!ReportToPrint
        Case 1
            DoCmd.OpenReport "Employee Sales by Country", PrintMode
        Case 2
            DoCmd.OpenReport "Sales Totals by Amount", PrintMode
        Case 3
            If IsNull(Forms![Sales Reports Dialog]!SelectCategory) Then
                DoCmd.OpenReport "Sales by Category", PrintMode
            Else
                DoCmd.OpenReport "Sales by Category", PrintMode, , strWhereCategory
            End If
    End Select
    DoCmd.Close acForm, "Sales Reports Dialog"

Exit_Preview_Click:
    Exit Sub

Err_Preview_Click:
    Resume Exit_Preview_Click

End Sub

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

Private Sub Preview_Click()
' Preview selected report. This procedure uses the PrintReports
' Sub procedure defined in (General) section of this module.
    
    PrintReports acPreview

End Sub




Private Sub Print_Click()
' Print selected report. This procedure uses the PrintReports
' Sub procedure defined in (General) section of this module.

    PrintReports acNormal
    
End Sub



Private Sub ReportToPrint_AfterUpdate()
' Enable SelectCategory combo box if user selected Sales by Category
' report.

    Const conSalesByCategory = 3

    If Me!ReportToPrint.Value = conSalesByCategory Then
        Me!SelectCategory.Enabled = True
    Else
        Me!SelectCategory.Enabled = False
    End If
    
End Sub



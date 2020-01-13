VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Employee Sales by Country"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
' Set page number to 1 when a new group starts.

    Page = 1
    
End Sub


Private Sub GroupHeader2_Format(Cancel As Integer, FormatCount As Integer)
' Display ExceededGoalLabel and SalespersonLine if salesperson's total
' meets criteria.

    If Me!SalespersonTotal > 5000 Then
        Me!ExceededGoalLabel.Visible = True
        Me!SalespersonLine.Visible = True
    Else
        Me!ExceededGoalLabel.Visible = False
        Me!SalespersonLine.Visible = False
    End If


End Sub


Private Sub Report_NoData(Cancel As Integer)
' Display a message if user enters a date for which there are no records,
' and don't preview or print report.

    Dim strMsg As String, strTitle As String
    Dim intStyle As Integer
    
    strMsg = "You must enter a date between 10-July-1996 and 06-May-1998."
    intStyle = vbOKOnly
    strTitle = "No Data for Date Range"
    
    MsgBox strMsg, intStyle, strTitle
    Cancel = True

End Sub



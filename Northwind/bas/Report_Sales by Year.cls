VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Sales by Year"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

' This variable set to True while the Report_Open event is executing.
Public blnOpening As Boolean
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
' If ShowDetails check box on Sales by Year Dialog form is cleared,
' don't print detail section.

    If Forms![Sales by Year Dialog]!ShowDetails = False Then Cancel = True

End Sub


Private Sub GroupFooter1_Format(Cancel As Integer, FormatCount As Integer)
    
    ' If ShowDetails check box on Sales by Year Dialog form is checked,
    ' set value of Show text box to False so that page header won't print
    ' on next page.
    If Forms![Sales by Year Dialog]!ShowDetails Then
        Me!Show.Value = False
    Else
    ' If ShowDetails check box on Sales by Year Dialog form is cleared,
    ' don't print group footer.
        Cancel = True
    End If
    
End Sub


Private Sub GroupFooter1_Retreat()
' If ShowDetails check box on Sales by Year Dialog form is checked,
' set value of Show text box to True so that page header will print on
' next page.

    If Forms![Sales by Year Dialog]!ShowDetails Then Me!Show.Value = True
    
End Sub


Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)

    Const conNone = 0
    
    ' If ShowDetails check box on Sales by Year Dialog form is checked,
    ' set value of Show text box to True so that page header will print on
    ' next page.
    If Forms![Sales by Year Dialog]!ShowDetails Then
        Me!Show.Value = True
    Else
    ' If ShowDetails check box on Sales by Year Dialog form is cleared, set
    ' ForceNewPage property to conNone and hide all labels for detail records.
        Me.Section(acGroupLevel1Footer).ForceNewPage = conNone
        Me!Line10.Visible = False
        Me!Line15.Visible = False
        Me!DetailsLabel.Visible = False
        Me!LineNumberLabel2.Visible = False
        Me!ShippedDateLabel2.Visible = False
        Me!OrdersShippedLabel2.Visible = False
        Me!SalesLabel2.Visible = False
    End If
    
End Sub

Private Sub PageHeader_Format(Cancel As Integer, FormatCount As Integer)
' If value of Show text box is False, don't print page header .

    If Me!Show.Value = False Then Cancel = True

End Sub


Private Sub Report_Close()
' Close the Sales by Year Dialog form.

    Dim strDocName As String

    strDocName = "Sales by Year Dialog"

    DoCmd.Close acForm, strDocName

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

Private Sub Report_Open(Cancel As Integer)
' Open Sales by Year Dialog form.
' IsLoaded function (defined in Utility Functions module) determines
' if specified form is open.
    
    Dim strDocName As String

    strDocName = "Sales by Year Dialog"
    ' Set public variable to True so Sales by Year Dialog knows that report
    ' is in its Open event.
    blnOpening = True
    
    ' Open form.
    DoCmd.OpenForm strDocName, , , , , acDialog
    
    ' If Sales by Year Dialog form isn't loaded, don't preview or print report.
    ' (User clicked Cancel button on form.)
    If IsLoaded(strDocName) = False Then Cancel = True
    
    'Set public variable to False, signifying that Open event is finished.
    blnOpening = False
End Sub


Private Sub ReportHeader_Format(Cancel As Integer, FormatCount As Integer)
' If ShowDetails check box is cleared, set Show text box to False.

    If Forms![Sales by Year Dialog]!ShowDetails = False Then Me!Show.Value = False
    
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rptEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error Resume Next
    Dim lngLen As Long
    lngLen = Len(Me.TotalPay)
    
    If lngLen > 4 Then
        ' Ricky's paycheck
        Me.TotalPay.ForeColor = vbRed
        Me.TotalPay.FontSize = 18
        Me.TotalPay.Top = 4 * 1440
        Me.TotalPay.Height = 1 * 1440
        Me.TotalPay.Width = 3 * 1440
    Else
        Me.TotalPay.ForeColor = vbBlack
        Me.TotalPay.FontSize = 8
        Me.TotalPay.Top = 3.8958 * 1440
        Me.TotalPay.Height = 0.1875 * 1440
        Me.TotalPay.Width = 1.5833 * 1440
    End If
        
End Sub

Private Sub Report_Close()
On Error Resume Next
    DoCmd.OpenForm "frmEmployees", , , , , acWindowNormal
End Sub

Private Sub Report_NoData(Cancel As Integer)
    Cancel = True
    MsgBox "No data to display."
End Sub

Private Sub Report_Open(Cancel As Integer)
On Error Resume Next
    DoCmd.OpenForm "frmEmployees", , , , , acHidden
End Sub

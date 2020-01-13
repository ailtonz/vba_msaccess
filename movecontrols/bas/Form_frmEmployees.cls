VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdReport_Click()
On Error GoTo Err_cmdReport_Click

    Dim stDocName As String

    stDocName = "rptEmployees"
    If Me.EmployeeID > 0 Then
        DoCmd.OpenReport stDocName, acPreview, , "[EmployeeID]=" & Me.EmployeeID
    Else
        MsgBox "No data to display."
    End If

Exit_cmdReport_Click:
    Exit Sub

Err_cmdReport_Click:
    MsgBox Err.Description
    Resume Exit_cmdReport_Click
    
End Sub

Private Sub Form_Current()
On Error Resume Next

    Dim lngLen As Long
    Dim bOK As Boolean
    
    If Me.Address.Value <> "" Then
        lngLen = Len(Me.Address.Value)
        If lngLen > 20 Then
            ' show larger address text box
            bOK = ToggleTextBox(True)
        Else
            ' show smaller address text box
            bOK = ToggleTextBox(False)
        End If
    Else
        ' no data
        bOK = ToggleTextBox(False)
    End If
    Debug.Print "Text box visibility and form size change" & IIf(bOK, "d successfully.", " failed.")
   
    
End Sub

Private Function ToggleTextBox(bLarge As Boolean) As Boolean
On Error Resume Next

    ToggleTextBox = True
    For Each ctl In Me.Controls
        If ctl.Tag = "largetext" Then
            ctl.Visible = bLarge
        ElseIf ctl.Tag = "smalltext" Then
            ctl.Visible = Not bLarge
        End If
    Next ctl
    
    If bLarge Then
        DoCmd.MoveSize , , , 3.95 * 1440
    Else
        DoCmd.MoveSize , , , 3.35 * 1440
    End If
            
    If Err <> 0 Then
        ToggleTextBox = True
    End If
End Function

Private Sub Form_Load()
On Error Resume Next

    Dim lngLen As Long
    Dim bOK As Boolean
    
    If Me.Address.Value <> "" Then
        lngLen = Len(Me.Address.Value)
        If lngLen > 20 Then
            ' show larger address text box
            bOK = ToggleTextBox(True)
        Else
            ' show smaller address text box
            bOK = ToggleTextBox(False)
        End If
        Debug.Print "Text box visibility change" & IIf(bOK, "d successfully.", " failed.")
    End If
End Sub

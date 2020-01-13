VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Customer Orders Subform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Sub Form_Current()
' This code created by Form Wizard.
    Dim strParentDocName As String

    On Error Resume Next
    strParentDocName = Me.Parent.Name

    If Err <> 0 Then
        GoTo Form_Current_Exit
    Else
        On Error GoTo Form_Current_Err
        Me.Parent![Customer Orders Subform2].Requery
    End If

Form_Current_Exit:
    Exit Sub

Form_Current_Err:
    MsgBox Err.Description
    Resume Form_Current_Exit

End Sub


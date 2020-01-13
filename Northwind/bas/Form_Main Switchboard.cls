VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main Switchboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database  ' Use database order for string comparisons.
Option Explicit  ' Requires variables to be declared before they are used.

Function OpenForms(strFormName As String) As Integer
' This function is used in the Click event of command buttons that
' open forms on the Main Switchboard. Using a function is more efficient
' than repeating the same code in multiple event procedures.
On Error GoTo Err_OpenForms

    ' Open specified form.
    DoCmd.OpenForm strFormName

Exit_OpenForms:
    Exit Function

Err_OpenForms:
    MsgBox Err.Description
    Resume Exit_OpenForms

End Function

Sub ExitMicrosoftAccess_Click()
' This code created by Command Button Wizard.
On Error GoTo Err_ExitMicrosoftAccess_Click

    ' Exit Microsoft Access.
    DoCmd.Quit

Exit_ExitMicrosoftAccess_Click:
    Exit Sub

Err_ExitMicrosoftAccess_Click:
    MsgBox Err.Description
    Resume Exit_ExitMicrosoftAccess_Click
    
End Sub
Sub DisplayDatabaseWindow_Click()
' This code created in part by Command Button Wizard.
On Error GoTo Err_DisplayDatabaseWindow_Click

    Dim strDocName As String
        
    strDocName = "Categories"

    ' Close Main Switchboard form.
    DoCmd.Close
    
    ' Give focus to Database window; select Categories table (first
    ' form in list).
    DoCmd.SelectObject acTable, strDocName, True

Exit_DisplayDatabaseWindow_Click:
    Exit Sub

Err_DisplayDatabaseWindow_Click:
    MsgBox Err.Description
    Resume Exit_DisplayDatabaseWindow_Click
    
End Sub






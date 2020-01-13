Attribute VB_Name = "Startup"
Option Compare Database
Option Explicit
' Functions in this module are used in the Startup form.

Function OpenStartup() As Boolean
' Displays Startup form only if database is not a design master or replica.
' Used in OnOpen property of Startup form.
On Error GoTo OpenStartup_Err
    If IsItAReplica() Then
        ' This database is a design master or replica, so close Startup form
        ' before it is displayed.
         DoCmd.Close
    Else
        ' This database is not a design master or replica, so display Startup form.
        ' Set the value of HideStartupForm check box using the value of
        ' StartupForm property of database (as set in code or in the
        ' Display Form/Page box in Startup dialog box).
        If (CurrentDb().Properties("StartupForm") = "Startup" Or _
            CurrentDb().Properties("StartupForm") = "Form.Startup") Then
            ' StartupForm property is set to Startup, so clear HideStartupForm
            ' check box.
            Forms!Startup!HideStartupForm = False
        Else
            ' StartupForm property is not set to Startup, so check HideStartupForm
            ' checkbox.
            Forms!Startup!HideStartupForm = True
        End If
    End If
   
OpenStartup_Exit:
    Exit Function
        
OpenStartup_Err:
    Const conPropertyNotFound = 3270
    If Err = conPropertyNotFound Then
        Forms!Startup!HideStartupForm = True
        Resume OpenStartup_Exit
    End If
End Function

Function HideStartupForm()
On Error GoTo HideStartupForm_Err
' Uses the value of HideStartupForm check box to determine the setting for
' StartupForm property of database. (The setting is displayed in Display Form
' box in Startup dialog box).
' Used in OnClose property of Startup form.
        If Forms!Startup!HideStartupForm Then
        ' HideStartupForm check box is checked, so set StartupForm property to Main SwitchBoard.
            CurrentDb().Properties("StartupForm") = "Main SwitchBoard"
        Else
            ' HideStartupForm check box is cleared, so set StartupForm property to Startup.
            CurrentDb().Properties("StartupForm") = "Startup"
        End If
    
        Exit Function
        
HideStartupForm_Err:
    Const conPropertyNotFound = 3270
    If Err = conPropertyNotFound Then
        Dim db As DAO.Database
        Dim prop As DAO.Property
        Set db = CurrentDb()
        Set prop = db.CreateProperty("StartupForm", dbText, "Startup")
        db.Properties.Append prop
        Resume Next
    End If
End Function
Function CloseForm()
' Closes Startup form.
' Used in OnClick property of OK command button on Startup form.
    DoCmd.Close
    DoCmd.OpenForm ("Main Switchboard")
End Function
Function IsItAReplica() As Boolean
On Error GoTo IsItAReplica_Err
' Determines if database is a design master or a replica.
' Used in OpenStartup function.
 
    Dim blnReturnValue As Boolean
    
    blnReturnValue = False
    If CurrentDb().Properties("Replicable") = "T" Then
        ' Replicable property setting is "T",
        ' so database is a design master or replica.
        blnReturnValue = True
    Else
        ' Replicable property setting is not "T",
        ' so database is not a design master or replica.
        blnReturnValue = False
    End If
    
IsItAReplica_Exit:
    IsItAReplica = blnReturnValue
    Exit Function
    
IsItAReplica_Err:
    Resume IsItAReplica_Exit

End Function


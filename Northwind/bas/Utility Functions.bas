Attribute VB_Name = "Utility Functions"
Option Compare Database
Option Explicit

Function IsLoaded(ByVal strFormName As String) As Boolean
 ' Returns True if the specified form is open in Form view or Datasheet view.
    Dim oAccessObject As AccessObject

    Set oAccessObject = CurrentProject.AllForms(strFormName)
    If oAccessObject.IsLoaded Then
        If oAccessObject.CurrentView <> acCurViewDesign Then
            IsLoaded = True
        End If
    End If
    
End Function




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim path As String

Private Sub AddPicture_Click()
    ' Use the Office File Open dialog to get a file name to use
    ' as an employee picture.
    getFileName
End Sub

Private Sub Form_RecordExit(Cancel As Integer)
    ' Hide the errormsg label to reduce flashing when navigating
    ' between records.
    errormsg.Visible = False
End Sub

Private Sub RemovePicture_Click()
    ' Clear the file name for the employee record and display the
    ' errormsg label.
    Me![ImagePath] = ""
    hideImageFrame
    errormsg.Visible = True
End Sub

Private Sub Form_AfterUpdate()
    ' Requery the ReportsTo combo box after a record has been changed.
    ' Then, either show the errormsg label if no file name exists for
    ' the employee record or display the image if there is a file name that
    ' exists.
    Me!ReportsTo.Requery
    On Error Resume Next
        showErrorMessage
        showImageFrame
        If (IsRelative(Me!ImagePath) = True) Then
            Me![ImageFrame].Picture = path & Me![ImagePath]
        Else
            Me![ImageFrame].Picture = Me![ImagePath]
        End If
End Sub

Private Sub ImagePath_AfterUpdate()
    ' After selecting an image for the employee, display it.
    On Error Resume Next
        showErrorMessage
        showImageFrame
        If (IsRelative(Me!ImagePath) = True) Then
            Me![ImageFrame].Picture = path & Me![ImagePath]
        Else
            Me![ImageFrame].Picture = Me![ImagePath]
        End If
End Sub
Private Sub Form_Current()
    ' Display the picture for the current employee record if the image
    ' exists.  If the file name no longer exists or the file name was blank
    ' for the current employee, set the errormsg label caption to the
    ' appropriate message.
    Dim res As Boolean
    Dim fName As String
    
    path = CurrentProject.path
    On Error Resume Next
        errormsg.Visible = False
        If Not IsNull(Me!Photo) Then
            res = IsRelative(Me!Photo)
            fName = Me![ImagePath]
            If (res = True) Then
                fName = path & "\" & fName
            End If
            
            Me![ImageFrame].Picture = fName
            showImageFrame
            Me.PaintPalette = Me![ImageFrame].ObjectPalette
            If (Me![ImageFrame].Picture <> fName) Then
                hideImageFrame
                errormsg.Caption = "Picture not found"
                errormsg.Visible = True
            End If
        Else
            hideImageFrame
            errormsg.Caption = "Click Add/Change to add picture"
            errormsg.Visible = True
        End If

End Sub

Sub getFileName()
    ' Displays the Office File Open dialog to choose a file name
    ' for the current employee record.  If the user selects a file
    ' display it in the image control.
    Dim fileName As String
    Dim result As Integer
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Employee Picture"
        .Filters.Add "All Files", "*.*"
        .Filters.Add "JPEGs", "*.jpg"
        .Filters.Add "Bitmaps", "*.bmp"
        .FilterIndex = 3
        .AllowMultiSelect = False
        .InitialFileName = CurrentProject.path
        result = .Show
        If (result <> 0) Then
            fileName = Trim(.SelectedItems.Item(1))
            Me![ImagePath].Visible = True
            Me![ImagePath].SetFocus
            Me![ImagePath].Text = fileName
            Me![FirstName].SetFocus
            Me![ImagePath].Visible = False
        End If
    End With
End Sub

Sub showErrorMessage()
    ' Display the errormsg label if the image file is not available.
    If Not IsNull(Me!Photo) Then
        errormsg.Visible = False
    Else
        errormsg.Visible = True
    End If
End Sub

Function IsRelative(fName As String) As Boolean
    ' Return false if the file name contains a drive or UNC path
    IsRelative = (InStr(1, fName, ":") = 0) And (InStr(1, fName, "\\") = 0)
End Function

Sub hideImageFrame()
    ' Hide the image control
    Me![ImageFrame].Visible = False
End Sub

Sub showImageFrame()
    ' Display the image control
    Me![ImageFrame].Visible = True
End Sub

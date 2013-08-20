VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkCreateProjectForm 
   Caption         =   "Create New Project"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "vtkCreateProjectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vtkCreateProjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : BrowseButton_Click
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Open a browse window and initialize the project folder text field
'---------------------------------------------------------------------------------------
'
Private Sub BrowseButton_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count > 0 Then
            ProjectPathTextBox.Text = .SelectedItems(1)
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreateButton_Click
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Create the project
'             If the Create Button is enabled, all parameters are OK
'---------------------------------------------------------------------------------------
'
Private Sub CreateButton_Click()
    vtkCreateProject path:=ProjectPathTextBox.Text, name:=ProjectNameTextBox.Text
    Unload vtkCreateProjectForm
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CancelButton_Click
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Quit the form when canceled
'---------------------------------------------------------------------------------------
'
Private Sub CancelButton_Click()
' don''t use End , because it stop debugger
    Unload vtkCreateProjectForm
End Sub

Private Sub ProjectPathTextBox_Change()
    enableCreateButton
End Sub

Private Sub ProjectNameTextBox_Change()
    enableCreateButton
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Deactivate the Create Button when creating the UserForm
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    enableCreateButton
End Sub

'---------------------------------------------------------------------------------------
' Procedure : enableCreateButton
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Enable the Create Button only if all parameters are typed and OK
'---------------------------------------------------------------------------------------
'
Private Sub enableCreateButton()
    Dim folderExists As Boolean, projectDoesntExists As Boolean, sep As String
    Dim fso As New FileSystemObject
    Const PINK = &HC0E0FF
    Const GREEN = &HC0FFC0
    
    ' Validate the path textField, the path must exist
   On Error Resume Next
    fso.GetFolder (ProjectPathTextBox.Text)
    folderExists = err.Number = 0
    If folderExists Then ProjectPathTextBox.BackColor = GREEN Else ProjectPathTextBox.BackColor = PINK
    
    ' Validate the Project name textField ; the project folder must not exist
    If Right$(ProjectPathTextBox.Text, 1) Like "\" Then sep = "" Else sep = "\"
    fso.GetFolder (ProjectPathTextBox.Text & sep & ProjectNameTextBox.Text) 'Will raise an error 76 if wrong path or not a folder
    projectDoesntExists = err.Number = 76
    If (Not folderExists And ProjectNameTextBox.Text Like "") Or (folderExists And Not projectDoesntExists) _
            Then ProjectNameTextBox.BackColor = PINK Else ProjectNameTextBox.BackColor = GREEN
   On Error GoTo 0
    
    ' Enable Create Button only if all parameters are OK
    CreateButton.Enabled = folderExists And projectDoesntExists
End Sub

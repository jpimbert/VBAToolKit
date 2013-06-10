VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkCreateProjectForm 
   Caption         =   "Create New Project"
   ClientHeight    =   2880
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
'Option Explicit

Private Sub CommandButton1_Click()
If TextBox1.Text <> "" And TextBox2.Text <> "" Then
retval = vtkCreateProject(TextBox2.Text, TextBox1.Text)
End
Else
MsgBox (" the two fields must be filled")
End If
End Sub

Private Sub CommandButton2_Click()

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
         
        If .SelectedItems.Count > 0 Then
            TextBox2.Text = .SelectedItems(1)
        End If
         
    End With
End Sub

Private Sub CommandButton3_Click()
' don''t use End , because it stop debugger
Unload vtkCreateProjectForm
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

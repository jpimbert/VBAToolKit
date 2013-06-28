VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VtkAddModule 
   Caption         =   "Create New Module"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "VtkAddModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VtkAddModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Option Explicit

Private Sub CommandButton1_Click()

If TextBox1.Text <> "" And (OptionButton1 <> "" Or OptionButton2 <> "" Or OptionButton3 <> "") Then


    If ActiveSheet.Cells(1, 1).Value <> "vtkConfigurations v1.0" Then
        MsgBox (" you can't run this command outside VbaToolkit Projects")
    ElseIf OptionButton1.Value = True Then
            retval = VtkAddOneModule(TextBox1.Text, 1)
        ElseIf OptionButton2.Value = True Then
            retval = VtkAddOneModule(TextBox1.Text, 2)
        ElseIf OptionButton3.Value = True Then
            retval = VtkAddOneModule(TextBox1.Text, 3)
    End If
Else
MsgBox ("the two fields must be filled")

End If
End Sub

Private Sub CommandButton2_Click()
' don''t use End , because it stop debugger
Unload VtkAddModule
End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub UserForm_Click()

End Sub

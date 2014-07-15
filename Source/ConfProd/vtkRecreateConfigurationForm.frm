VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkRecreateConfigurationForm 
   Caption         =   "Recreate Configuration"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   OleObjectBlob   =   "vtkRecreateConfigurationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vtkRecreateConfigurationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkCreateProjectForm
' Author    : Jean-Pierre IMBERT
' Purpose   : UserForm for VBAToolKit configuration recreation
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------


Private m_confManager As vtkConfigurationManager
Private m_XMLFilePath As String
Private m_XMLFileOK As Boolean
Private m_ConfSelected As Boolean

Private Const PINK = &HC0E0FF
Private Const GREEN = &HC0FFC0


'---------------------------------------------------------------------------------------
' Procedure : BrowseButton_Click
' Author    : Jean-Pierre Imbert
' Date      : 15/07/2014
' Purpose   : Open a browse window and initialize the XML Configuration File Text field
'---------------------------------------------------------------------------------------
'
Private Sub BrowseButton_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count > 0 Then
            XMLFileTextBox.Text = .SelectedItems(1)
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Lucas Vitorino
' Purpose   : Initializing global variables.
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    
    ' Select the previous XML File Path if any
    XMLFileTextBox.Text = m_XMLFilePath
    validateXMLFileTextBox
    
    ' Disable the "Create Configuration" button as no configuration is selected
    enableReCreateButton

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConfigurationListBox_AfterUpdate
' Author    : Jean-Pierre Imbert
' Purpose   : Manage the list box containing the list of configurations
'---------------------------------------------------------------------------------------
'
Private Sub ConfigurationListBox_AfterUpdate()
    If m_ConfSelected Then ConfigurationListBox.BackColor = GREEN Else ConfigurationListBox.BackColor = PINK
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConfigurationListBox_Change
' Author    : Jean-Pierre Imbert
' Purpose   : Manage the list box containing the list of configurations
'---------------------------------------------------------------------------------------
'
Private Sub ConfigurationListBox_Change()
    Dim i As Integer
    m_ConfSelected = False
    For i = 0 To ConfigurationListBox.ListCount - 1
        m_ConfSelected = m_ConfSelected Or ConfigurationListBox.Selected(i)
    Next i
    enableReCreateButton
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CancelButton_Click
' Author    : Lucas Vitorino
' Purpose   : Close the form
'---------------------------------------------------------------------------------------
'
Private Sub CancelButton_Click()
    vtkRecreateConfigurationForm.Hide
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateConfigurationButton_Click
' Author    : Lucas Vitorino
' Purpose   : Call the vtkRecreateConfigurations Sub with the relevant parameters
'---------------------------------------------------------------------------------------
'
Private Sub CreateConfigurationButton_Click()
    Me.Hide
    ' display wait message modeless for the present code to keep running
    vtkWaitForm.Show vbModeless
    ' build the confNames collection
    Dim confNames As New Collection, i As Integer
    For i = 0 To ConfigurationListBox.ListCount - 1
        If ConfigurationListBox.Selected(i) Then confNames.Add ConfigurationListBox.List(i)
    Next i
    ' recreate configurations
    vtkRecreateConfigurations m_confManager, confNames
    ' Hide wait message
    vtkWaitForm.Hide
End Sub

'---------------------------------------------------------------------------------------
' Procedure : validateXMLFileTextBox
' Author    : Jean-Pierre Imbert
' Date      : 15/07/2014
' Purpose   : Check the XMLFileTextBox and establish status of the form
'
'---------------------------------------------------------------------------------------
'
Private Sub validateXMLFileTextBox()
    Dim cmX As New vtkConfigurationManagerXML, conf As vtkConfiguration
   On Error Resume Next
    cmX.init XMLFileTextBox.Text
    m_XMLFileOK = (Err.Number = 0)
   On Error GoTo 0
    ConfigurationListBox.Clear
    m_ConfSelected = False
    If m_XMLFileOK Then
        XMLFileTextBox.BackColor = GREEN
        m_XMLFilePath = XMLFileTextBox.Text
        Set m_confManager = cmX
        For Each conf In m_confManager.configurations
            ConfigurationListBox.AddItem conf.name
        Next conf
       Else
        XMLFileTextBox.BackColor = PINK
        m_XMLFilePath = ""
        Set m_confManager = Nothing
    End If
    enableReCreateButton
End Sub

'---------------------------------------------------------------------------------------
' Procedure : enableReCreateButton
' Author    : Jean-Pierre Imbert
' Date      : 15/07/2014
' Purpose   : Enable the ReCreate Button only if all parameters are typed and OK
'
'---------------------------------------------------------------------------------------
'
Private Sub enableReCreateButton()
    ' Enable ReCreate Button only if all parameters are OK
    CreateConfigurationButton.Enabled = m_XMLFileOK And m_ConfSelected
End Sub


Private Sub XMLFileTextBox_Change()
    validateXMLFileTextBox
End Sub

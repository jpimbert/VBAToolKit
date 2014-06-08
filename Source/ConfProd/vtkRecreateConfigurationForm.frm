VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkRecreateConfigurationForm 
   Caption         =   "Recreate Configuration"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "vtkRecreateConfigurationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vtkRecreateConfigurationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Module    : vtkCreateProjectForm
' Author    : Lucas Vitorino
' Purpose   : UserForm for VBAToolKit configuration recreation
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
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


Private cm As vtkConfigurationManager
Private currentConf As vtkConfiguration
Private currentProjectName As String

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Lucas Vitorino
' Purpose   : Initializing global variables.
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()

    ' Get the name of the current DEV workbook
    currentProjectName = getCurrentProjectName

    ' Initialize configuration manager
    Set cm = vtkConfigurationManagerForProject(currentProjectName)
    
    ' Disable the "Create Configuration" button as no configuration is selected
    enableCreateConfigurationButton

    ' Initialize the content of the combo box
    If Not cm Is Nothing Then
        Dim conf As vtkConfiguration
        For Each conf In cm.configurations
            ConfigurationComboBox.AddItem (conf.name)
        Next
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConfigurationComboBox_Change
' Author    : Lucas Vitorino
' Purpose   : Manage the combo box containing the list of configurations
'---------------------------------------------------------------------------------------
'
Private Sub ConfigurationComboBox_Change()
    
    On Error GoTo ConfigurationComboBox_Change_Error

    Set currentConf = cm.configurations(ConfigurationComboBox.Value)
    
    If AllConfigurationsExceptThisOneCheckBox.Value = False Then
        PathTextBox.Text = currentConf.path
    End If
    
    enableCreateConfigurationButton

    On Error GoTo 0
    Exit Sub

ConfigurationComboBox_Change_Error:
    Set currentConf = Nothing
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AllConfigurationsExceptThisOneCheckBox_Change
' Author    : Lucas Vitorino
' Purpose   : Manage what happens when the checkbox changes.
'---------------------------------------------------------------------------------------
'
Private Sub AllConfigurationsExceptThisOneCheckBox_Change()

    If AllConfigurationsExceptThisOneCheckBox.Value = True Or currentConf Is Nothing Then
        PathTextBox.Text = ""
    Else
        PathTextBox.Text = currentConf.path
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CancelButton_Click
' Author    : Lucas Vitorino
' Purpose   : Close the form
'---------------------------------------------------------------------------------------
'
Private Sub CancelButton_Click()
    Unload vtkRecreateConfigurationForm
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateConfigurationButton_Click
' Author    : Lucas Vitorino
' Purpose   : Call the vtkRecreateConfiguration function with the relevant parameters
'---------------------------------------------------------------------------------------
'
Private Sub CreateConfigurationButton_Click()
    
    On Error GoTo CreateConfigurationButton_Click_Error
    
    If AllConfigurationsExceptThisOneCheckBox.Value = False Then
        vtkRecreateConfiguration currentProjectName, currentConf.name
    Else
        Dim conf As vtkConfiguration
        For Each conf In cm.configurations
            If Not conf.name Like currentConf.name Then
                vtkRecreateConfiguration currentProjectName, conf.name
            End If
        Next
    End If
    
    On Error GoTo 0
    Exit Sub

CreateConfigurationButton_Click_Error:
    Err.Source = "CreateConfigurationButton_Click of module vtkRecreateConfigurationForm"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_ALREADY_OPEN ' Trying to replace an open workbook while recreating the configuration
            MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
            Resume Next
        Case VTK_NO_SOURCE_FILES ' A source file is missing
            MsgBox "Error " & Err.Number & " (" & Err.Description & ")"
            Resume Next
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
    
    Exit Sub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : enableCreateConfigurationButton
' Author    : Lucas Vitorino
' Purpose   : Decide if the "Create Configuration" button should be enabled or disabled.
'---------------------------------------------------------------------------------------
'
Private Sub enableCreateConfigurationButton()
    If currentConf Is Nothing Then
        CreateConfigurationButton.Enabled = False
    Else
        CreateConfigurationButton.Enabled = True
    End If
End Sub

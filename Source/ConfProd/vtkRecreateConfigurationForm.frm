VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkRecreateConfigurationForm 
   Caption         =   "Recreate Configration"
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

    ' Temporary
    ' TODO : implement way of keeping track of the current project
    currentProjectName = "VBAToolKit"

    ' Initialize configuration manager
    Set cm = vtkConfigurationManagerForProject(currentProjectName)
    
    ' Disable the "Create Configuration" button as no configuration is selected
    enableCreateConfigurationButton

    ' Initialize the content of the combo box
    Dim conf As vtkConfiguration
    For Each conf In cm.configurations
        ConfigurationComboBox.AddItem (conf.name)
    Next
    
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
    PathTextBox.Text = currentConf.path
    
    enableCreateConfigurationButton

    On Error GoTo 0
    Exit Sub

ConfigurationComboBox_Change_Error:
    Set currentConf = Nothing
    Resume Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CancelButton_Click
' Author    : Lucas Vitorino
' Purpose   : Close the form
'---------------------------------------------------------------------------------------
'
Private Sub CancelButton_Click()
    Unload VBAToolKit_DEV.vtkRecreateConfigurationForm
End Sub



'---------------------------------------------------------------------------------------
' Procedure : CreateConfigurationButton_Click
' Author    : Lucas Vitorino
' Purpose   : Call the vtkRecreateConfiguration function with the relevant parameters
'---------------------------------------------------------------------------------------
'
Private Sub CreateConfigurationButton_Click()
    vtkRecreateConfiguration currentProjectName, currentConf.name
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

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vtkNewRecreateConfigurationForm 
   Caption         =   "VBAToolKit - Recreate Configuration"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   OleObjectBlob   =   "vtkNewRecreateConfigurationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vtkNewRecreateConfigurationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkNewRecreateConfigurationForm
' Author    : Lucas Vitorino
' Purpose   : UserForm for VBAToolKit Recreate Configuration feature
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

Private fso As New FileSystemObject

'---------------------------------------------------------------------------------------
' Procedure : UserForm_Initialize
' Author    : Lucas Vitorino
' Purpose   : Initialize the different objects and variables in the form
'---------------------------------------------------------------------------------------
'
Private Sub UserForm_Initialize()
    
    Dim dummyDom As New MSXML2.DOMDocument
    
    ' Manage the 'list of projects' objects
    ListOfProjectsTextBox.Text = xmlRememberedProjectsFullPath
    
    ' Ugly with bool1 and bool2 but the one liner condition does not work in my VM
    Dim bool1 As Boolean: bool1 = fso.FileExists(xmlRememberedProjectsFullPath)
    Dim bool2 As Boolean: bool2 = dummyDom.Load(xmlRememberedProjectsFullPath)
    If bool1 And bool2 Then
        ' Everything is fine
        ListOfProjectsTextBox.ForeColor = &HC000&
        
        Dim tmpProj As New vtkProject
        For Each tmpProj In listOfRememberedProjects
            ListOfProjectsComboBox.AddItem tmpProj.projectName
        Next
        
    Else
        ListOfProjectsTextBox.ForeColor = &HFF&
    End If
    

End Sub

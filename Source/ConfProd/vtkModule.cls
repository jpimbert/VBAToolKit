VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkModule
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Contains informations relative to a VTK Module
'               - 'name' of the module (wrapper to 'parent' methods)
'               - 'parent' is the configuration manager in which is included the module
'               - 'number' (read only) is the number of the module in the configuration manager
'               - 'ID' (read only) is the ID of the module for XML exportation, build with the number
'             If the module is attached to a configuration at its initialization
'               - 'configuration' (read only) if the module is attached to a configuration (wrapper to 'parent' methods)
'               - 'path' (read/write) is the full path for module code source for the configuration (wrapper to 'parent' methods)
'               - 'VBAModule' (read) VBA code module of the configuration if the Excel file is open
'             A vtkModule object must be initialized with 'init' method
'             Methods :
'               - getPathForConfiguration
'               - setPathForConfiguration
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

Private m_parent As vtkConfigurationManager             ' A configuration is attached to a Configuration Manager
Private m_number As Integer
Private m_confNumber As Integer

'---------------------------------------------------------------------------------------
' Procedure : init
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Initialize a vtkConfiguration object with parent and number
'
' NOTE      : This method is dedicated to be used by vtkConfigurationManager
'---------------------------------------------------------------------------------------
'
Public Sub init(confManager As vtkConfigurationManager, modNumber As Integer, Optional confNumber As Integer = 0)
    Set m_parent = confManager
    m_number = modNumber
    m_confNumber = confNumber
End Sub

'---------------------------------------------------------------------------------------
' Property  : parent as vtkConfigurationManager (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The vtkConfiguration in which is declared this vtkModule
'---------------------------------------------------------------------------------------
'
Public Property Get parent() As vtkConfigurationManager
    Set parent = m_parent
End Property

'---------------------------------------------------------------------------------------
' Property  : number as Integer (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The number of the module in its parent
'---------------------------------------------------------------------------------------
'
Public Property Get Number() As Integer
    Number = m_number
End Property

'---------------------------------------------------------------------------------------
' Property  : ID as String (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The ID of the module for XML exportation
'---------------------------------------------------------------------------------------
'
Public Property Get ID() As String
    ID = "m" & m_number
End Property

'---------------------------------------------------------------------------------------
' Property  : confNumber as Integer (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The number in its parent of the configuration attached to the module
'             or 0 if no configuration is attached
'---------------------------------------------------------------------------------------
'
Public Property Get confNumber() As Integer
    confNumber = m_confNumber
End Property

'---------------------------------------------------------------------------------------
' Property  : name as string (Read Only, wrapper on the parent vtkConfigurationManager)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The name of the module
'---------------------------------------------------------------------------------------
'
Property Get name() As String
    name = parent.module(Number)
End Property

'---------------------------------------------------------------------------------------
' Procedure : getPathForConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : get the path of the module for a specific configuration
' Error Raised :
'           VTK_INEXISTANT_CONFIGURATION if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getPathForConfiguration(confName As String) As String
    Dim cn As Integer
    cn = parent.getConfigurationNumber(confName)
    getPathForConfiguration = ""
    If cn = 0 Then Err.Raise VTK_INEXISTANT_CONFIGURATION, "vtkModule getPathForConfiguration", "Configuration """ & confName & """ in project """ & parent.projectName & """"
    getPathForConfiguration = parent.getModulePathWithNumber(Number, cn)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setPathForConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : change the path of the module for a specific configuration
' Error Raised :
'           VTK_INEXISTANT_CONFIGURATION if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function setPathForConfiguration(confName As String, newPath As String) As String
    Dim cn As Integer
    cn = parent.getConfigurationNumber(confName)
    If cn = 0 Then Err.Raise VTK_INEXISTANT_CONFIGURATION, "vtkModule getPathForConfiguration", "Configuration """ & confName & """ in project """ & parent.projectName & """"
    parent.setModulePathWithNumber path:=newPath, numModule:=Number, numConfiguration:=cn
End Function

'---------------------------------------------------------------------------------------
' Procedure : configuration
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Get the name of the configuration attached to the module
' Error Raised :
'           VTK_MODULE_NOTATTACHED if the module is not attached to a configuration
'---------------------------------------------------------------------------------------
'
Public Property Get configuration() As String
    configuration = ""
    If confNumber = 0 Then Err.Raise VTK_MODULE_NOTATTACHED, "vtkModule::configuration()", "Module """ & parent.module(Number) & """ not attached to a configuration "
    configuration = parent.configuration(confNumber)
End Property

'---------------------------------------------------------------------------------------
' Procedure : path
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Get the path of the module in the configuration attached
' Error Raised :
'           VTK_MODULE_NOTATTACHED if the module is not attached to a configuration
'---------------------------------------------------------------------------------------
'
Public Property Get path() As String
    path = ""
    If confNumber = 0 Then Err.Raise VTK_MODULE_NOTATTACHED, "vtkModule::path()", "Module """ & parent.module(Number) & """ not attached to a configuration"
    path = parent.getModulePathWithNumber(numModule:=Number, numConfiguration:=confNumber)
End Property

'---------------------------------------------------------------------------------------
' Procedure : VBAModule
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Get the VBComponent of the module in the configuration attached
' Error Raised :
'           VTK_MODULE_NOTATTACHED if the module is not attached to a configuration
'           VTK_WORKBOOK_NOTOPEN if the Workbook of the configuration is not open
'           no error raised if Component is not in the configuration, just return Nothing
'---------------------------------------------------------------------------------------
'
Public Property Get VBAModule() As VBComponent
    Dim Wb As Workbook, fso As New FileSystemObject, wbName As String
    Set VBAModule = Nothing
    If confNumber = 0 Then Err.Raise VTK_MODULE_NOTATTACHED, "vtkModule::VBAModule()", "Module """ & parent.module(Number) & """ not attached to a configuration"
    wbName = fso.GetFileName(parent.getConfigurationPathWithNumber(confNumber))
    If Not VtkWorkbookIsOpen(wbName) Then Err.Raise VTK_WORKBOOK_NOTOPEN, "vtkModule::configuration()", "Workbook """ & wbName & """ is not open"
   On Error Resume Next
    Set VBAModule = Workbooks(wbName).VBProject.VBComponents(name)
   On Error GoTo 0
End Property

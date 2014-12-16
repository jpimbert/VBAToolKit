VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Interface : vtkConfigurationManager
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Interface for Configurations, References and Modules management of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a Module can be imported/exported to a specific path relative to the VTK Project directory
'               - each configuration is a VBA project included in a specific Excel file whose path is relative to the VTK Project directory
'               - for each configuration, the set of VBA references is managed
'               - some other parameters are managed for each configration
' NOTE      : This interface is adopted by Excel and XML versions of a configuration manager
'               - vtkConfigurationManagerExcel and vtkConfigurationManagerXML
'             Only the getter methods are concerned by the interface, for now only the Excel version
'             contains setter methods for the configuration parameters
'
' Usage:
'   - Each instance of Configuration Manager is attached to a VTK project)
'       - the method vtkConfigurationManagerForProject of module vtkConfigurationManagers give the instance attached to a project, or create it
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

' init (avec XMLPath spécifique pour XML, avant d'avoir la gestion centralisée des projets)

' rootPath
Public Property Get rootPath() As String
' projectName (uniquement en Get pour la première version avant gestion des projets)
Public Property Get projectName() As String
Attribute projectName.VB_UserMemId = 0
End Property

' configurationCount
Public Property Get configurationCount() As Integer
' configurations
Public Function configurations() As Collection
' configuration
Public Function configuration(n As Integer) As String
' getConfigurationNumber
Public Function getConfigurationNumber(configuration As String) As Integer
' setConfigurationPath
Public Sub setConfigurationPath(configuration As String, path As String)
' getConfigurationPath
Public Function getConfigurationPath(configuration As String) As String
' getConfigurationPathWithNumber
Public Function getConfigurationPathWithNumber(n As Integer) As String
' setConfigurationPathWithNumber <- (dans vtkConfiguration, ne sont utilisés que pour les tests)
Public Sub setConfigurationPathWithNumber(n As Integer, path As String)
' getConfigurationProjectNameWithNumber
Public Function getConfigurationProjectNameWithNumber(n As Integer) As String
' setConfigurationProjectNameWithNumber <-
Public Sub setConfigurationProjectNameWithNumber(n As Integer, projectName As String)
' getConfigurationCommentWithNumber
Public Function getConfigurationCommentWithNumber(n As Integer) As String
' setConfigurationCommentWithNumber <-
Public Sub setConfigurationCommentWithNumber(n As Integer, comment As String)
' getConfigurationTemplateWithNumber
Public Function getConfigurationTemplateWithNumber(n As Integer) As String
' setConfigurationTemplateWithNumber <-
Public Sub setConfigurationTemplateWithNumber(n As Integer, template As String)
' getConfigurationPasswordWithNumber
Public Function getConfigurationPasswordWithNumber(n As Integer) As String
' setConfigurationPasswordWithNumber <-
Public Sub setConfigurationPasswordWithNumber(n As Integer, password As String)

End Sub
' addConfiguration
Public Function addConfiguration(configuration As String, Optional path As String = "", Optional template As String = "", Optional projectName As String = "", Optional comment As String = "") As Integer
' --> on garde les setters dans l'interface, avec une erreur dans XML

' moduleCount
Public Property Get moduleCount() As Integer
' modules
Public Function modules() As Collection
' module
Public Function module(n As Integer) As String
' getModuleNumber
Public Function getModuleNumber(module As String) As Integer
' getModulePathWithNumber
Public Function getModulePathWithNumber(numModule As Integer, numConfiguration As Integer) As String
' setModulePathWithNumber <- (dans vtkModule, non utilisé à part dans les tests)
Public Sub setModulePathWithNumber(path As String, numModule As Integer, numConfiguration As Integer)
' addModule
Public Function addModule(module As String) As Integer

' references
Public Function references() As Collection
' getConfigurationReferencesWithNumber
Public Function getConfigurationReferencesWithNumber(n As Integer) As Collection


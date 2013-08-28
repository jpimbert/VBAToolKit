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
Property Get parent() As vtkConfigurationManager
    Set parent = m_parent
End Property

'---------------------------------------------------------------------------------------
' Property  : number as Integer (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The number of the module in its parent
'---------------------------------------------------------------------------------------
'
Property Get number() As Integer
    number = m_number
End Property

'---------------------------------------------------------------------------------------
' Property  : name as string (Read Only, wrapper on the parent vtkConfigurationManager)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The name of the module
'---------------------------------------------------------------------------------------
'
Property Get name() As String
    name = m_parent.module(m_number)
End Property


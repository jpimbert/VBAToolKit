VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Contains informations relative to a VTK Configuration
'               - 'name' of the configuration (wrapper to 'parent' methods)
'               - 'parent' is the configuration manager in which is included the configuration
'               - 'number' (read only) is the number of the configuration in the configuration manager
'               - 'path' is the full path for the Excel workbook containing the configuration (wrapper to 'parent' methods)
'               - 'modules' is a collection of vtkModule objects, all the modules contained in the configuration with non null path (wrapper to 'parent' methods)
'             A vtkConfiguration object must be initialized with 'init' method
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

'---------------------------------------------------------------------------------------
' Procedure : init
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Initialize a vtkConfiguration object with parent and number
'
' NOTE      : This method is dedicated to be used by vtkConfigurationManager
'---------------------------------------------------------------------------------------
'
Public Sub init(confManager As vtkConfigurationManager, confNumber As Integer)
    Set m_parent = confManager
    m_number = confNumber
End Sub

'---------------------------------------------------------------------------------------
' Property  : parent as vtkConfigurationManager (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The vtkConfiguration in which is declared this vtkConfiguration
'---------------------------------------------------------------------------------------
'
Property Get parent() As vtkConfigurationManager
    Set parent = m_parent
End Property

'---------------------------------------------------------------------------------------
' Property  : number as Integer (Read Only, initialized with init)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The number of the configuration in its parent
'---------------------------------------------------------------------------------------
'
Property Get number() As Integer
    number = m_number
End Property

'---------------------------------------------------------------------------------------
' Property  : name as string (Read Only, wrapper on the parent vtkConfigurationManager)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The name of the configuration
'---------------------------------------------------------------------------------------
'
Property Get name() As String
    name = m_parent.configuration(m_number)
End Property

'---------------------------------------------------------------------------------------
' Property  : path as string (wrapper on the parent vtkConfigurationManager)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : The relative path of the configuration
'---------------------------------------------------------------------------------------
'
Property Get path() As String
    path = m_parent.getConfigurationPathWithNumber(m_number)
End Property

Property Let path(newPath As String)
    m_parent.setConfigurationPathWithNumber m_number, newPath
End Property

'---------------------------------------------------------------------------------------
' Property  : modules as collection (Read Only)
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : All modules with non null path are included in the collection
'             the vtkModules are attached to the configuration
'---------------------------------------------------------------------------------------
'
Property Get modules() As Collection
    Dim col As New Collection, i As Integer, M As vtkModule
    For i = 1 To m_parent.moduleCount
        If Not (m_parent.getModulePathWithNumber(numModule:=i, numConfiguration:=m_number) Like "") Then
            Set M = New vtkModule
            M.init confManager:=m_parent, modNumber:=i, confNumber:=m_number
            col.Add Item:=M, Key:=m_parent.module(i)
        End If
    Next i
    Set modules = col
End Property

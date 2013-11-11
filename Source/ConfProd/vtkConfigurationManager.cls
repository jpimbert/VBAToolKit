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
' Module    : vtkConfigurationManager
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Manage Configurations and Modules of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a Module can be imported/exported to a specific path relative to the VTK Project directory
'               - each configuration is a VBA project included in a specific Excel file whose path is relative to the VTK Project directory
'
' Usage:
'   - Each instance of Configuration Manager is attached to the DEV Excel Worksheet of a VTK project)
'       - the method vtkConfigurationManagerForProject of module vtkConfigurationManagers give the instance attached to a project, or create it
'   - a VTK project is implemented by a main Excel file containing a sheet named "vtkConfigurations"
'       - if such a sheet doesn't exist, it is created
'   - When initializing an instance, the relative development project workbook must be opened
'       - else initialization is not performed
'
' To implement later
'   - automatic recognition of the configuration sheet version, and proposal to the user to convert to the last version
'       cell $A$1 contains the version identification of the sheet version like "vtkConfigurations v1.0"
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

Private m_projectName As String             ' The configuration of a project is attached to the DEV Excel file
Private m_workbook As Workbook
Private m_configurationSheet As Worksheet
Private m_rootPath As String
Private Const sheetName As String = "vtkConfigurations"
Private Const currentVersion As String = "vtkConfigurations v2.0"
Private Const nbTitleColumns As Integer = 1
Private Const nbTitleRows As Integer = 2

Private fso As New FileSystemObject
Private m_xmlSheetPath As String
Private m_configurationDom As MSXML2.DOMDocument

Private Sub Class_Terminate()
    Set m_configurationSheet = Nothing
    Set m_workbook = Nothing
    
    Set m_configurationDom = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Property  : projectName as String - Read/Write - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : name of the project attached to this configuration manager
'               - when changed, the configurationWorkSheet private property is seset
' NOTE      : This is the Default property
'               - "Attribute Value.VB_UserMemId = 0" is added as first line of property Get (in exported Module)
' WARNING   : The attribute must be explicitly added after each export, or just before import
'---------------------------------------------------------------------------------------
'
Public Property Get projectName() As String
Attribute projectName.VB_UserMemId = 0
    projectName = m_projectName
End Property

Public Property Let projectName(ByVal sprojectName As String)

    Dim sheetpath As String
    sheetpath = fso.BuildPath(vtkRootPathForProject(sprojectName), vtkXmlRelPathForProject(sprojectName))
    If m_projectName <> sprojectName And fso.FileExists(sheetpath) Then
        m_projectName = sprojectName
        m_xmlSheetPath = sheetpath
        m_rootPath = ""
        'setConfigurationSheet      ' Reset the configuration sheet when the attached project is changed
    End If
    
End Property

'---------------------------------------------------------------------------------------
' Property  : rootPath as String - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : return the rootPath of the project if the Excel file was already saved
'             - or return "" if the rootPath can't be determined
'             - the rootPath is cached when identified the first time
'---------------------------------------------------------------------------------------
'
Public Property Get rootPath() As String
    Dim ret As String
    If Not m_rootPath Like "" Then
        ret = m_rootPath
       Else
        ret = ""
        ret = vtkRootPathForProject(m_projectName)
        If Not ret Like "" Then
            m_rootPath = ret
        End If
    End If
    
    rootPath = ret

End Property


'---------------------------------------------------------------------------------------
' Procedure : configurationSheet as Worksheet - Read only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the configuration Sheet - used during tests
'---------------------------------------------------------------------------------------
'
Public Property Get configurationSheet() As Worksheet
    Set configurationSheet = m_configurationSheet
End Property


'---------------------------------------------------------------------------------------
' Procedure : configurationDOM
' Author    : Lucas Vitorino
' Purpose   : Return the XML DOM with all the information - used during tests
'---------------------------------------------------------------------------------------
'
Public Property Get configurationDOM() As MSXML2.DOMDocument
    
    On Error GoTo configurationDOM_Error

    Dim dom As New MSXML2.DOMDocument
    dom.async = False
    dom.validateOnParse = True
    If dom.Load(m_xmlSheetPath) <> True Then
        Set configurationDOM = Nothing
        Err.Raise VTK_SHEET_NOT_VALID
        Exit Property
    End If
    Set configurationDOM = dom

    On Error GoTo 0
    Exit Property

configurationDOM_Error:
    Err.Source = "vtkConfigurationManager::configurationDOM"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            Err.Description = "The xml sheet for project " & projectName & " has failed to load."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
            
    Exit Property

End Property

'---------------------------------------------------------------------------------------
' Property  : configurationCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Get the number of configurations
'---------------------------------------------------------------------------------------
'
Public Property Get configurationCount() As Integer
    ' The number of configurations is the number of used columns minus one
    On Error GoTo configurationCount_Error

    configurationCount = configurationDOM.SelectNodes("/vtkConf/configuration").Length

    On Error GoTo 0
    Exit Property

configurationCount_Error:
    Err.Source = "vtkConfigurationManager::configurationCount"

    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but forward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Property

'---------------------------------------------------------------------------------------
' Function  : configuration
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : return the n-th configuration of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function configuration(n As Integer) As String
    
    On Error GoTo configuration_Error
    
    Dim sResult As String
    sResult = ""

    If n >= 1 And n <= configurationCount Then
        sResult = configurationDOM.SelectSingleNode("/vtkConf/configuration[" & n - 1 & "]/name[0]").Text
    Else
        sResult = ""
    End If
    
    configuration = sResult

    On Error GoTo 0
    Exit Function

configuration_Error:
    Err.Source = "vtkConfigurationManager::configuration"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Procedure : configurations
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Return a collection of initialized vtkConfiguration objects
'---------------------------------------------------------------------------------------
'
Public Function configurations() As Collection
    
    On Error GoTo configurations_Error

    Dim col As New Collection, i As Integer, c As vtkConfiguration
    
    For i = 1 To configurationCount
        Set c = New vtkConfiguration
        c.init confManager:=Me, confNumber:=i
        col.Add Item:=C, Key:=configuration(i)
    Next i
    Set configurations = col

    On Error GoTo 0
    Exit Function

configurations_Error:
    Err.Source = "vtkConfigurationManager::configurations"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationNumber
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the number of a configuration given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationNumber(configuration As String) As Integer

    Dim i As Integer
    On Error GoTo getConfigurationNumber_Error

    For i = 1 To configurationCount
        If configurationDOM.SelectSingleNode("/vtkConf/configuration[" & i - 1 & "]/name[0]").Text = configuration Then
            getConfigurationNumber = i
            Exit Function
        End If
    Next
    
    getConfigurationNumber = 0

    On Error GoTo 0
    Exit Function

getConfigurationNumber_Error:
    Err.Source = "vtkConfigurationManager::getConfigurationNumber"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationPathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : return the n-th relative configuration path of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPathWithNumber(n As Integer) As String

    On Error GoTo getConfigurationPathWithNumber_Error
    
    Dim sResult As String
    sResult = ""
    
    If n >= 1 And n <= configurationCount Then
        sResult = configurationDOM.SelectSingleNode("/vtkConf/configuration[" & n - 1 & "]/path[0]").Text
       Else
        sResult = ""
    End If
    
    getConfigurationPathWithNumber = sResult

    On Error GoTo 0
    Exit Function

getConfigurationPathWithNumber_Error:
    Err.Source = "vtkConfigurationManager::getConfigurationPathWithNumber"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Function  : getConfigurationPath
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the relative path of a configuration given its name
'               - return "" of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPath(configuration As String) As String
    On Error GoTo getConfigurationPath_Error

    getConfigurationPath = getConfigurationPathWithNumber(getConfigurationNumber(configuration))

    On Error GoTo 0
    Exit Function

getConfigurationPath_Error:
    Err.Source = "vtkConfigurationManager::getConfigurationPath"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Property  : moduleCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Get the number of modules
'---------------------------------------------------------------------------------------
'
Public Property Get moduleCount() As Integer
    ' The number of modules is the number of used rows minus two
    On Error GoTo moduleCount_Error

    moduleCount = configurationDOM.SelectNodes("/vtkConf/module").Length

    On Error GoTo 0
    Exit Property

moduleCount_Error:
    Err.Source = "vtkConfigurationManager::moduleCount"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Property

End Property

'---------------------------------------------------------------------------------------
' Procedure : modules
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Return a collection of initialized vtkModule objects
'---------------------------------------------------------------------------------------
'
Public Function modules() As Collection

    On Error GoTo modules_Error

    Dim col As New Collection, i As Integer, m As vtkModule
    
    For i = 1 To moduleCount
        Set m = New vtkModule
        m.init confManager:=Me, modNumber:=i
        col.Add Item:=m, Key:=module(i)
    Next i
    
    Set modules = col

    On Error GoTo 0
    Exit Function

modules_Error:
    Err.Source = "vtkConfigurationManager::modules"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Function  : module
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the n-th module of the project, given n as integer
'               - return "" if the module is inexistant
'---------------------------------------------------------------------------------------
'
Public Function module(n As Integer) As String
    
    On Error GoTo module_Error
    
    Dim sResult As String
    sResult = ""
    If n >= 1 And n <= moduleCount Then
        sResult = configurationDOM.SelectSingleNode("/vtkConf/module[" & n - 1 & "]/name[0]").Text
       Else
        sResult = ""
    End If
    module = sResult

    On Error GoTo 0
    Exit Function

module_Error:
    Err.Source = "vtkConfigurationManager::module"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function

End Function

'---------------------------------------------------------------------------------------
' Procedure : getModuleNumber
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the number of a module given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getModuleNumber(module As String) As Integer
    
    On Error GoTo getModuleNumber_Error
    
    Dim i As Integer
    For i = 1 To moduleCount
        If configurationDOM.SelectSingleNode("/vtkConf/module[" & i - 1 & "]/name[0]").Text = module Then
            getModuleNumber = i
            Exit Function
        End If
    Next
    getModuleNumber = 0

    On Error GoTo 0
    Exit Function

getModuleNumber_Error:
    Err.Source = "vtkConfigurationManager::getModuleNumber"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function
    
End Function

'''---------------------------------------------------------------------------------------
''' Procedure : addModule
''' Author    : Jean-Pierre Imbert
''' Date      : 26/05/2013
''' Purpose   : Add a new module given its name
'''               - return the number of the added module
'''               - if the module is already existant, return minus the number of the existant module
'''---------------------------------------------------------------------------------------
'''
''Public Function addModule(module As String) As Integer
''    Dim n As Integer
''    n = getModuleNumber(module)
''    If n = 0 Then   ' n = 0 means non-existant module
''        n = moduleCount + 1                                     ' Set module number
''        configurationSheet.Cells(n + nbTitleRows, 1) = module   ' Write module name
''        Dim i As Integer
''        For i = 1 To configurationCount                         ' For each configuration
''            configurationSheet.Cells(n + nbTitleRows, i + nbTitleColumns) = "-"     ' Set Default path as not initialized
''        Next i
''        addModule = n
''       Else
''        addModule = -n
''    End If
''End Function

'---------------------------------------------------------------------------------------
' Procedure : getModulePathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : Get a module relative path given its module and configuration numbers
'               - return "" if the module or the configuration doesn't exist
'               - return "" if the path is not initialized ("-" in the cell of the configuration sheet)
'---------------------------------------------------------------------------------------
'
Public Function getModulePathWithNumber(numModule As Integer, numConfiguration As Integer) As String

    On Error GoTo getModulePathWithNumber_Error
    
    Dim sResult As String
    sResult = ""
    If numModule >= 1 And numModule <= moduleCount And numConfiguration >= 1 And numConfiguration <= configurationCount Then
        'sResult = configurationSheet.Cells(numModule + nbTitleRows, numConfiguration + nbTitleColumns)
        'If sResult = "-" Then sResult = ""
        On Error Resume Next
        sResult = configurationDOM.SelectSingleNode("/vtkConf/module[" & numModule - 1 & "]/modulePath[@confId=""" & _
                    configurationDOM.SelectSingleNode("/vtkConf/configuration[" & numConfiguration - 1 & "]/@cID").Text & _
                    """]").Text
        On Error GoTo 0
       Else
        sResult = ""
    End If
    getModulePathWithNumber = sResult

    On Error GoTo 0
    Exit Function

getModulePathWithNumber_Error:
    Err.Source = "vtkConfigurationManager::getModulePathWithNumber"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Function


End Function


'---------------------------------------------------------------------------------------
' Procedure : refCount
' Author    : Lucas Vitorino
' Purpose   : Get the number of references used by all the configurations of a project.
'---------------------------------------------------------------------------------------
'
Public Property Get refCount()

    On Error GoTo refCount_Error

    refCount = configurationDOM.SelectNodes("/vtkConf/reference").Length

    On Error GoTo 0
    Exit Property

refCount_Error:
    Err.Source = "vtkConfigurationManager::refCount"
    
    Select Case Err.Number
        Case VTK_SHEET_NOT_VALID
            ' do nothing but froward the error
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Property

End Property


''---------------------------------------------------------------------------------------
'' Procedure : getReferencesForConfiguration
'' Author    : Lucas Vitorino
'' Purpose   : Get the collection of references for the n-th configuration.
'' Return    : A collection of VBIDE.Reference objects.
''---------------------------------------------------------------------------------------
''
'Public Function getReferencesForConfiguration(numConfiguration As Integer) As Collection
'
'    On Error GoTo getReferencesForConfiguration_Error
'
'    Dim retCol As New Collection
'    Dim tmpNode As MSXML2.IXMLDOMNode
'    Dim tmpRef As VBIDE.Reference
'    Dim confId As Integer
'
'    If numConfiguration >= 1 And numConfiguration <= refCount Then
'        confId = configurationDOM.SelectSingleNode("vtkConf/configuration[" & numConfiguration - 1 & "]/@cID").Text
'
'        For Each tmpNode In configurationDOM.SelectNodes("vtkConf/reference[@confIDs[contains(.,""" & confId & """)]]")
'            tmpRef.name = tmpNode.SelectSingleNode("/name[0]").Text
'
'            ' The DTD ensures both the fields will not be filled
'            If configurationDOM.SelectNodes("/guid").Length > 0 Then
'                'tmpRef.GUID = configurationDOM.SelectSingleNode("/vtkConf/reference[" & i - 1 & "]/guid[0]").Text
'                tmpRef.GUID = configurationDOM.SelectSingleNode("/guid[0]").Text
'            End If
'
'            If configurationDOM.SelectNodes("/path").Length > 0 Then
'                'tmpRref.fullPath = configurationDOM.SelectSingleNode("/vtkConf/reference[" & i - 1 & "]/path[0]").Text
'                tmpRref.fullPath = configurationDOM.SelectSingleNode("/path[0]").Text
'            End If
'
'            col.Add Item:=tmpRef, Key:=tmpRef.name
'
'            ' Reinitialize the reference object
'            tmpRef.name = ""
'            tmpRef.GUID = ""
'            tmpRref.fullPath = ""
'
'        Next
'
'    End If
'
'    On Error GoTo 0
'    Set getReferencesForConfiguration = retCol
'    Exit Function
'
'getReferencesForConfiguration_Error:
'    Err.Source = "vtkConfigurationManager::getReferencesForConfiguration"
'
'    Select Case Err.Number
'        Case VTK_SHEET_NOT_VALID
'            ' do nothing but froward the error
'        Case Else
'            Err.Number = VTK_UNEXPECTED_ERROR
'    End Select
'
'    Err.Raise Err.Number, Err.Source, Err.Description
'
'    Exit Function
'
'End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManagerXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements vtkConfigurationManager
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagerXML, implements vtkConfigurationManager
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Manage Configurations and Modules of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a Module can be imported/exported to a specific path relative to the VTK Project directory
'               - each configuration is a VBA project included in a specific Excel file whose path is relative to the VTK Project directory
'
' Usage:
'   - Each instance of Configuration Manager is attached to the XML file describing a VTK project
'       - the method init is used for this attachment
'
' NOTE      : The XML configuration manager is read only
'             a call to a setter function will cause an error
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

Private m_DOMDocument As DOMDocument
Private m_projectName As String
Private m_rootPath As String

Private Sub Class_Terminate()
    Set m_DOMDocument = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Sub       : init - Public
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Attach the instance to an XML file, and initialize
' Parameter : XMLFilePath (absolute path) as a string
' Error detected :
'           - VTK_WRONG_FILE_PATH if the XML file can't be opened
'           - VTK_INVALID_XML_FILE if the XML File is not a V2.0 Configuration file
'           - VTK_OBSOLETE_CONFIGURATION_SHEET if the XML file is not on the good version
'           In case of error the instance keeps the previous DOM document
'---------------------------------------------------------------------------------------
'
Public Sub init(ByVal XMLFilePath As String)
    Dim fso As New FileSystemObject, DDoc As MSXML2.DOMDocument
    Dim projectName As String, rootPath As String
    If fso.FileExists(XMLFilePath) Then
        Set DDoc = New MSXML2.DOMDocument
        DDoc.Load XMLFilePath
        DDoc.setProperty "SelectionLanguage", "XPath"  ' for the use of the contains() function of XPath
        If DDoc.parseError.ErrorCode <> 0 Then
            Err.Raise Number:=VTK_INVALID_XML_FILE, Source:="init:vtkConfigurationManagerXML", Description:="Parse error of " & XMLFilePath & " : " & DDoc.parseError.reason
        End If
        If DDoc.SelectSingleNode("/vtkConf/info/vtkConfigurationsVersion").Text <> "2.0" Then
            Err.Raise Number:=VTK_OBSOLETE_CONFIGURATION_SHEET, Source:="init:vtkConfigurationManagerXML", Description:="Bad version of XML File : 2.0 expected."
        End If
       On Error GoTo M_Error
        projectName = DDoc.SelectSingleNode("/vtkConf/info/projectName").Text
        rootPath = fso.GetParentFolderName(fso.GetParentFolderName(XMLFilePath))
       On Error GoTo 0
        Set m_DOMDocument = DDoc
        m_projectName = projectName
        m_rootPath = rootPath
       Else
        Err.Raise Number:=VTK_WRONG_FILE_PATH, Source:="init:vtkConfigurationManagerXML", Description:="The File " & XMLFilePath & " is unreachable."
    End If
    Exit Sub
M_Error:
    Err.Raise Number:=VTK_INVALID_XML_FILE, Source:="init:vtkConfigurationManagerXML", Description:="Parse error of " & XMLFilePath & " : " & DDoc.parseError.reason
End Sub

'---------------------------------------------------------------------------------------
' Property  : projectName as String - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : name of the project attached to this configuration manager
' NOTE      : This is the Default property
'               - "Attribute Value.VB_UserMemId = 0" is added as first line of property Get (in exported Module)
' WARNING   : The attribute must be explicitly added after each export, or just before import
'---------------------------------------------------------------------------------------
'
Public Property Get projectName() As String
    projectName = m_projectName
End Property

Public Property Get vtkConfigurationManager_projectName() As String
    vtkConfigurationManager_projectName = projectName
End Property

'---------------------------------------------------------------------------------------
' Property  : rootPath as String - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the rootPath of the project
'             - or return "" if the rootPath can't be determined
'---------------------------------------------------------------------------------------
'
Public Property Get rootPath() As String
    rootPath = m_rootPath
End Property

Public Property Get vtkConfigurationManager_rootPath() As String
    vtkConfigurationManager_rootPath = rootPath
End Property

'---------------------------------------------------------------------------------------
' Property  : configurationCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Get the number of configurations
'---------------------------------------------------------------------------------------
'
Public Property Get configurationCount() As Integer
    configurationCount = m_DOMDocument.SelectNodes("/vtkConf/configuration").Length
End Property

Public Property Get vtkConfigurationManager_configurationCount() As Integer
    vtkConfigurationManager_configurationCount = configurationCount
End Property

'---------------------------------------------------------------------------------------
' Function  : configuration
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th configuration of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function configuration(n As Integer) As String
    If n > 0 And n <= configurationCount Then
        configuration = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/name").Text
       Else
        configuration = ""
    End If
End Function

Public Function vtkConfigurationManager_configuration(n As Integer) As String
    vtkConfigurationManager_configuration = configuration(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : raiseReadOnlyError
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Private Sub raiseReadOnlyError(Source As String)
    Err.Raise Number:=VTK_READONLY_FILE, Source:=Source, Description:="A XML Configuration File is considered Read-Only"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : addConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Function addConfiguration(configuration As String, Optional path As String = "", Optional template As String = "", Optional projectName As String = "", Optional comment As String = "") As Integer
    raiseReadOnlyError Source:="addConfiguration:vtkConfigurationManagerXML"
End Function

Public Function vtkConfigurationManager_addConfiguration(configuration As String, Optional path As String = "", Optional template As String = "", Optional projectName As String = "", Optional comment As String = "") As Integer
    vtkConfigurationManager_addConfiguration = addConfiguration(configuration, path, template, projectName, comment)
End Function

'---------------------------------------------------------------------------------------
' Procedure : configurations
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return a collection of initialized vtkConfiguration objects
'---------------------------------------------------------------------------------------
'
Public Function configurations() As Collection
    Dim col As New Collection, i As Integer, c As vtkConfiguration
    For i = 1 To configurationCount
        Set c = New vtkConfiguration
        c.init confManager:=Me, confNumber:=i
        col.Add Item:=c, Key:=configuration(i)
    Next i
    Set configurations = col
End Function

Public Function vtkConfigurationManager_configurations() As Collection
    Set vtkConfigurationManager_configurations = configurations
End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the number of a configuration given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationNumber(configuration As String) As Integer
    Dim textValue As String
   On Error Resume Next
    textValue = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[name=""" & configuration & """]/@cID").Text
   On Error GoTo 0
    If textValue = "" Then
        getConfigurationNumber = 0
       Else
        getConfigurationNumber = CInt(Right$(textValue, Len(textValue) - 1))
    End If
End Function

Public Function vtkConfigurationManager_getConfigurationNumber(configuration As String) As Integer
    vtkConfigurationManager_getConfigurationNumber = getConfigurationNumber(configuration)
End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationPathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th relative configuration path of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPathWithNumber(n As Integer) As String
    If n > 0 And n <= configurationCount Then
        getConfigurationPathWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/path").Text
       Else
        getConfigurationPathWithNumber = ""
    End If
End Function

Public Function vtkConfigurationManager_getConfigurationPathWithNumber(n As Integer) As String
    vtkConfigurationManager_getConfigurationPathWithNumber = getConfigurationPathWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationPathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationPathWithNumber(n As Integer, path As String)
    raiseReadOnlyError Source:="setConfigurationPathWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationPathWithNumber(n As Integer, path As String)
    setConfigurationPathWithNumber n, path
End Sub

'---------------------------------------------------------------------------------------
' Function  : getConfigurationPath
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the relative path of a configuration given its name
'               - return "" of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPath(configuration As String) As String
    getConfigurationPath = getConfigurationPathWithNumber(getConfigurationNumber(configuration))
End Function

Public Function vtkConfigurationManager_getConfigurationPath(configuration As String) As String
    vtkConfigurationManager_getConfigurationPath = getConfigurationPath(configuration)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationPath
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationPath(configuration As String, path As String)
    raiseReadOnlyError Source:="setConfigurationPath:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationPath(configuration As String, path As String)
    setConfigurationPath configuration, path
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationProjectNameWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th configuration project name of the project, given n as integer
'               - return "" if the configuration is inexistant
'               - return the default project name if
'                   - it is not initialized in the configuration sheet v1.1
'                   - the configuration sheet is v1.0, so there is no project name field
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationProjectNameWithNumber(n As Integer) As String
    If n > 0 And n <= configurationCount Then
        getConfigurationProjectNameWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/title").Text
       Else
        getConfigurationProjectNameWithNumber = ""
    End If
End Function

Public Function vtkConfigurationManager_getConfigurationProjectNameWithNumber(n As Integer) As String
    vtkConfigurationManager_getConfigurationProjectNameWithNumber = getConfigurationProjectNameWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationProjectNameWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationProjectNameWithNumber(n As Integer, projectName As String)
    raiseReadOnlyError Source:="setConfigurationProjectNameWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationProjectNameWithNumber(n As Integer, projectName As String)
    setConfigurationProjectNameWithNumber n, projectName
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationCommentWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th configuration comment of the project, given n as integer
'               - return "" if the configuration is inexistant
'               - return the default comment if
'                   - it is not initialized in the configuration sheet v1.1
'                   - the configuration sheet is v1.0, so there is no comment field
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationCommentWithNumber(n As Integer) As String
    If n > 0 And n <= configurationCount Then
        getConfigurationCommentWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/comment").Text
       Else
        getConfigurationCommentWithNumber = ""
    End If
End Function

Public Function vtkConfigurationManager_getConfigurationCommentWithNumber(n As Integer) As String
    vtkConfigurationManager_getConfigurationCommentWithNumber = getConfigurationCommentWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationCommentWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationCommentWithNumber(n As Integer, comment As String)
    raiseReadOnlyError Source:="setConfigurationCommentWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationCommentWithNumber(n As Integer, comment As String)
    setConfigurationCommentWithNumber n, comment
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationTemplateWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th configuration Excel template path of the project, given n as integer
'               - return "" if the configuration is inexistant
'               - return the default comment if
'                   - it is not initialized in the configuration sheet v1.1
'                   - the configuration sheet is v1.0, so there is no comment field
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationTemplateWithNumber(n As Integer) As String
    If n > 0 And n <= configurationCount Then
       On Error Resume Next
        getConfigurationTemplateWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/templatePath").Text
       On Error GoTo 0
       Else
        getConfigurationTemplateWithNumber = ""
    End If
End Function

Public Function vtkConfigurationManager_getConfigurationTemplateWithNumber(n As Integer) As String
    vtkConfigurationManager_getConfigurationTemplateWithNumber = getConfigurationTemplateWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationTemplateWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationTemplateWithNumber(n As Integer, template As String)
    raiseReadOnlyError Source:="setConfigurationTemplateWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationTemplateWithNumber(n As Integer, template As String)
    setConfigurationTemplateWithNumber n, template
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationPasswordWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 04/12/2014
' Purpose   : return the n-th configuration Excel code password of the project, given n as integer
'               - return "" if the configuration is inexistant
'               - return the default password if
'                   - it is not initialized in the configuration sheet v1.2
'                   - the configuration sheet is v1.0 or v1.2, so there is no password field
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPasswordWithNumber(n As Integer) As String
'    If n > 0 And n <= configurationCount Then
'       On Error Resume Next
'        getConfigurationPasswordWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]/templatePath").Text
'       On Error GoTo 0
'       Else
        getConfigurationPasswordWithNumber = ""
'    End If
End Function

Public Function vtkConfigurationManager_getConfigurationPasswordWithNumber(n As Integer) As String
    vtkConfigurationManager_getConfigurationPasswordWithNumber = getConfigurationPasswordWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationPasswordWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 04/12/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationPasswordWithNumber(n As Integer, password As String)
    raiseReadOnlyError Source:="setConfigurationPasswordWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setConfigurationPasswordWithNumber(n As Integer, password As String)
    setConfigurationPasswordWithNumber n, password
End Sub

'---------------------------------------------------------------------------------------
' Function  : vtkReferenceWithXML, private
' Author    : Jean-Pierre Imbert
' Date      : 08/07/2014
' Purpose   : Create and init a vtkReference with XML, given an Id
'---------------------------------------------------------------------------------------
'
Private Function vtkReferenceWithXML(id As String) As vtkReference
    Dim ref As New vtkReference
    ref.id = id
    ref.name = m_DOMDocument.SelectSingleNode("/vtkConf/reference[@refID=""" & id & """]/name").Text
   On Error Resume Next
    ref.GUID = m_DOMDocument.SelectSingleNode("/vtkConf/reference[@refID=""" & id & """]/guid").Text
    If Err.Number <> 0 Then
        Err.Number = 0
        ref.fullPath = m_DOMDocument.SelectSingleNode("/vtkConf/reference[@refID=""" & id & """]/path").Text
    End If
   On Error GoTo 0
    Set vtkReferenceWithXML = ref
End Function
'---------------------------------------------------------------------------------------
' Procedure : getConfigurationReferencesWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th configuration references collection of the project, given n as integer
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationReferencesWithNumber(n As Integer) As Collection
    Dim refIDs() As String, i As Integer, refs As New Collection
   On Error GoTo M_Error
    refIDs = Split(m_DOMDocument.SelectSingleNode("/vtkConf/configuration[@cID=""c" & n & """]").Attributes.getNamedItem("refIDs").Text)
   On Error GoTo 0
    For i = LBound(refIDs) To UBound(refIDs)
        refs.Add vtkReferenceWithXML(refIDs(i)), Key:=refIDs(i)
    Next i
M_Error:    ' refIDs attribute is optional
    Set getConfigurationReferencesWithNumber = refs
End Function

Public Function vtkConfigurationManager_getConfigurationReferencesWithNumber(n As Integer) As Collection
    Set vtkConfigurationManager_getConfigurationReferencesWithNumber = getConfigurationReferencesWithNumber(n)
End Function

'---------------------------------------------------------------------------------------
' Property  : moduleCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Get the number of modules
'---------------------------------------------------------------------------------------
'
Public Property Get moduleCount() As Integer
    moduleCount = m_DOMDocument.SelectNodes("/vtkConf/module").Length
End Property

Public Property Get vtkConfigurationManager_moduleCount() As Integer
    vtkConfigurationManager_moduleCount = moduleCount()
End Property

'---------------------------------------------------------------------------------------
' Procedure : modules
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return a collection of initialized vtkModule objects
'---------------------------------------------------------------------------------------
'
Public Function modules() As Collection
    Dim col As New Collection, i As Integer, m As vtkModule
    For i = 1 To moduleCount
        Set m = New vtkModule
        m.init confManager:=Me, modNumber:=i
        col.Add Item:=m, Key:=module(i)
    Next i
    Set modules = col
End Function

Public Function vtkConfigurationManager_modules() As Collection
    Set vtkConfigurationManager_modules = modules
End Function

'---------------------------------------------------------------------------------------
' Function  : module
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the n-th module of the project, given n as integer
'               - return "" if the module is inexistant
'---------------------------------------------------------------------------------------
'
Public Function module(n As Integer) As String
    If n > 0 And n <= moduleCount Then
        module = m_DOMDocument.SelectSingleNode("/vtkConf/module[@mID=""m" & n & """]/name").Text
       Else
        module = ""
    End If
End Function

Public Function vtkConfigurationManager_module(n As Integer) As String
    vtkConfigurationManager_module = module(n)
End Function

'---------------------------------------------------------------------------------------
' Procedure : getModuleNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : return the number of a module given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getModuleNumber(module As String) As Integer
    Dim textValue As String
   On Error Resume Next
    textValue = m_DOMDocument.SelectSingleNode("/vtkConf/module[name=""" & module & """]/@mID").Text
   On Error GoTo 0
    If textValue = "" Then
        getModuleNumber = 0
       Else
        getModuleNumber = CInt(Right$(textValue, Len(textValue) - 1))
    End If
End Function

Public Function vtkConfigurationManager_getModuleNumber(module As String) As Integer
    vtkConfigurationManager_getModuleNumber = getModuleNumber(module)
End Function

'---------------------------------------------------------------------------------------
' Procedure : addModule
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return the Error VTK_READONLY_FILE because XML file is unmodifiable
'---------------------------------------------------------------------------------------
'
Public Function addModule(module As String) As Integer
    raiseReadOnlyError Source:="addModule:vtkConfigurationManagerXML"
End Function

Public Function vtkConfigurationManager_addModule(module As String) As Integer
    vtkConfigurationManager_addModule = addModule(module)
End Function

'---------------------------------------------------------------------------------------
' Procedure : getModulePathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Get a module relative path given its module and configuration numbers
'               - return "" if the module or the configuration doesn't exist
'               - return "" if the path is not initialized ("-" in the cell of the configuration sheet)
'---------------------------------------------------------------------------------------
'
Public Function getModulePathWithNumber(numModule As Integer, numConfiguration As Integer) As String
    If numModule > 0 And numModule <= moduleCount And numConfiguration > 0 And numConfiguration <= configurationCount Then
       On Error Resume Next
        getModulePathWithNumber = m_DOMDocument.SelectSingleNode("/vtkConf/module[@mID=""m" & numModule & """]/modulePath[@confId=""c" & numConfiguration & """]").Text
       On Error GoTo 0
       Else
        getModulePathWithNumber = ""
    End If
End Function

Public Function vtkConfigurationManager_getModulePathWithNumber(numModule As Integer, numConfiguration As Integer) As String
    vtkConfigurationManager_getModulePathWithNumber = getModulePathWithNumber(numModule, numConfiguration)
End Function

'---------------------------------------------------------------------------------------
' Procedure : setModulePathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Set a module relative path given its module and configuration numbers
'               - if the path given is "", it is replaced by "-" in the configuration sheet cell
'---------------------------------------------------------------------------------------
'
Public Sub setModulePathWithNumber(path As String, numModule As Integer, numConfiguration As Integer)
    raiseReadOnlyError Source:="setModulePathWithNumber:vtkConfigurationManagerXML"
End Sub

Public Sub vtkConfigurationManager_setModulePathWithNumber(path As String, numModule As Integer, numConfiguration As Integer)
    setModulePathWithNumber path, numModule, numConfiguration
End Sub

'---------------------------------------------------------------------------------------
' Procedure : references
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Return a collection of all vtkReferences objects
'---------------------------------------------------------------------------------------
'
Public Function references() As Collection
    Dim i As Integer, refs As New Collection, id As String
    For i = 1 To m_DOMDocument.SelectNodes("/vtkConf/reference").Length
        id = "r" & i
        refs.Add vtkReferenceWithXML(id), Key:=id
    Next i
    Set references = refs
End Function

Public Function vtkConfigurationManager_references() As Collection
    Set vtkConfigurationManager_references = references
End Function


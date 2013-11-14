Attribute VB_Name = "vtkConfigurationManagers"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManagers
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Manage the configuration managers (class vtkConfigurationManager) for open projects
'
' Usage:
'   - Each instance of Configuration Manager is attached to the DEV Excel Workbook of a project
'       - the method vtkConfigurationManagerForProject give the instance attached to a workbook, or create it
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

'   collection of instances indexed by project names
Private configurationManagers As Collection

'   Combo type instead of completing the vtkConfiguration class
Public Type ConfWB
    conf As vtkConfiguration
    Wb As Workbook
    wasOpened As Boolean
End Type

'---------------------------------------------------------------------------------------
' Procedure : vtkConfigurationManagerForProject
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the configuration manager attached to the DEV Excel file given its project name
'               - if the configuration doesn't exist, it is created
'               - if the configurationManagers collection doesn't exist, it is created
'---------------------------------------------------------------------------------------
'
Public Function vtkConfigurationManagerForProject(projectName As String) As vtkConfigurationManager
    ' Create the collection if it doesn't exist
    If configurationManagers Is Nothing Then
        Set configurationManagers = New Collection
        End If
    ' search for the configuration manager in the collection
    Dim cm As vtkConfigurationManager
    On Error Resume Next
    Set cm = configurationManagers(projectName)
    If Err <> 0 Then
        Set cm = New vtkConfigurationManager
        cm.projectName = projectName
        If cm.projectName Like projectName Then     ' The initialization could fail (if the Workbook is closed)
            configurationManagers.Add Item:=cm, Key:=projectName
           Else
            Set cm = Nothing
        End If
    End If
   On Error GoTo 0
    ' return the configuration manager
    Set vtkConfigurationManagerForProject = cm
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkResetConfigurationManagers
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Reset all configuration managers (used during tests)
'---------------------------------------------------------------------------------------
'
Public Sub vtkResetConfigurationManagers()
    Set configurationManagers = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeConfigurationForActiveWorkBook
' Author    : Jean-Pierre Imbert
' Date      : 07/08/2013
'
' WARNING 1 : for now used only with manual run to convert a VBA project for VBAToolkit
' WARNING 2 : A beforeSave event handler is added even if one is already existing
'
' Purpose   : Create and Initialize a vtkConfiguration sheet for the active workbook
'             - does nothing if the active workbook already contains a vtkConfiguration worksheet
'             - initialize the worksheet with all VBA modules contained in the workbook
'             - a BeforeSave event handler is added to the new ActiveWorkbook
'             - manage VBAUnit, Tester class and standard modules appropriately
'             - the suffix "_DEV" is appended to the project name
'             - the Excel workbook is saved as a new file with DEV appended to the name
'             - the Delivery version is described in configuration but not created
'---------------------------------------------------------------------------------------
'
Public Sub vtkInitializeConfigurationForActiveWorkBook()
    ' If a configuration sheet exists, does nothing
    Dim cm As New vtkConfigurationManager
    If cm.isConfigurationInitializedForWorkbook(ExcelName:=ActiveWorkbook.name) Then Exit Sub
    Set cm = Nothing
    
    ' Get the project name and initialize a vtkProject with it
    Dim project As vtkProject
    Set project = vtkProjectForName(projectName:=ActiveWorkbook.VBProject.name)
    
    ' Change the project name
    ActiveWorkbook.VBProject.name = project.projectDEVName
    
    ' Change the workbook name
    ActiveWorkbook.SaveAs fileName:=ActiveWorkbook.path & "\" & project.workbookDEVName
    
    ' Prepare configuration manager
    Dim i As Integer, c As VBComponent, cn_dev As Integer, cn_prod As Integer, nm As Integer
    Set cm = vtkConfigurationManagerForProject(projectName:=project.projectName)
    cn_dev = cm.getConfigurationNumber(configuration:=project.projectDEVName)
    cn_prod = cm.getConfigurationNumber(configuration:=project.projectName)
    
    ' List all modules
    For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
        Set c = ActiveWorkbook.VBProject.VBComponents.Item(i)
        If c.Type <> vbext_ct_Document Then
            nm = cm.addModule(c.name)
            cm.setModulePathWithNumber path:=vtkStandardPathForModule(module:=c), numModule:=nm, numConfiguration:=cn_dev
            If vtkStandardCategoryForModuleName(moduleName:=c.name) Like "Prod" Then
                cm.setModulePathWithNumber path:=vtkStandardPathForModule(module:=c), numModule:=nm, numConfiguration:=cn_prod
            End If
        End If
    Next
    
    ' Add a BeforeSave event handler for the workbook
    vtkAddBeforeSaveHandlerInDEVWorkbook Wb:=ActiveWorkbook, projectName:=project.projectName, confName:=project.projectDEVName
    
    ' Save the new workbook
    ActiveWorkbook.Save
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkVerifyConfigurations
' Author    : Jean-Pierre IMBERT
' Date      : 13/11/2013
' Purpose   : Verify the coherence of the configurations description and the real
'             configuration workbooks and source modules
'             The verifications performed are :
'             - All configuration pathes are reachable
'             - All modules listed in a configuration description are existing in the configuration
'             - All modules really present in a configuration are described in the description with non null path
'   To be implemented in next commits
'             - All modules pathes are reachable
'   To be implemented perhaps (?)
'             - Each code module implemented in a configuration is the same as the source code module
'   To be implemented later (with XML configuration files)
'             - All references listed in a configuration description are existing in the configuration
'             - All references really present in a configuration are described in the description
'---------------------------------------------------------------------------------------
'
Sub vtkVerifyConfigurations()
    ' Init project and configuration manager
    Dim prj As vtkProject
    Set prj = vtkProjectForName(getCurrentProjectName)
    Dim cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(prj.projectName)
    
    ' Declare variables
    Dim c As vtkConfiguration, s As String, fso As New FileSystemObject
    Dim cwb() As ConfWB
    ReDim cwb(1 To cm.configurationCount) As ConfWB
    Debug.Print "----------------------------------------------------"
    Debug.Print "  Start verification of " & getCurrentProjectName & " project configurations"
    Debug.Print "----------------------------------------------------"
   
    ' Verify configuration pathes
    Dim nbConf As Integer
    nbConf = 0
    For Each c In cm.configurations
        s = cm.rootPath & "\" & c.path
        If fso.FileExists(s) Then
            nbConf = nbConf + 1
            Set cwb(nbConf).conf = c
           On Error Resume Next
            Set cwb(nbConf).Wb = Workbooks(fso.GetFileName(s))
           On Error GoTo 0
            cwb(nbConf).wasOpened = Not (cwb(nbConf).Wb Is Nothing)
            If Not cwb(nbConf).wasOpened Then Set cwb(nbConf).Wb = Workbooks.Open(fileName:=s, ReadOnly:=True)
           Else
            Debug.Print "Path of configuration " & c.name & " unreachable (" & s & ")."
        End If
    Next
    
    ' Verify that all modules in a configuration are in the description
    Dim i As Integer, mods As Collection, vbc As VBIDE.VBComponent, md As vtkModule
    For i = 1 To nbConf
        Set mods = cwb(i).conf.modules
        For Each vbc In cwb(i).Wb.VBProject.VBComponents
           On Error Resume Next
            Set md = mods(vbc.name)
            If Err.Number <> 0 Then
                Debug.Print "Module " & vbc.name & " is in configuration workbook " & cwb(i).conf.name & " but not in description of configuration."
            End If
           On Error GoTo 0
        Next
    Next i
    
    ' Verify that all modules in a description are in the configuration
    For i = 1 To nbConf
        For Each md In cwb(i).conf.modules
           On Error Resume Next
            Set vbc = cwb(i).Wb.VBProject.VBComponents(md.name)
            If Err.Number <> 0 Then
                Debug.Print "Module " & md.name & " is in the configuration description of " & cwb(i).conf.name & " but not in the workbook."
            End If
           On Error GoTo 0
        Next
    Next i
    
    ' Close all Worbooks opened during this verification
    For i = 1 To nbConf
        If Not cwb(i).wasOpened Then cwb(i).Wb.Close saveChanges:=False
    Next i
    Debug.Print "----------------------------------------------------"
End Sub

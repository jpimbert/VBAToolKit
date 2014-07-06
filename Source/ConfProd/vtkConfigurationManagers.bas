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
' NOTE      : For now this module uses only Excel Configuration Managers
'             The use of XML configuration managers needs a centralized Project management
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
' WARNING 3 : This function must use an Excel Configuration Manager (not XML)
'
' Purpose   : Create and Initialize a vtkConfiguration sheet for the active workbook
'             - does nothing if the active workbook already contains a vtkConfiguration worksheet
'             - initialize the worksheet with all VBA modules contained in the workbook
'             - a BeforeSave event handler is added to the new ActiveWorkbook
'             - manage VBAUnit, Tester class and standard modules appropriately
'             - the suffix "_DEV" is appended to the project name
'             - the Excel workbook is saved as a new file with DEV appended to the name
'             - the Delivery version is described in configuration but not created
'             - the reference sheet is created and initialized according to the actual references
'---------------------------------------------------------------------------------------
'
Public Sub vtkInitializeConfigurationForActiveWorkBook(Optional withBeforeSaveHandler As Boolean = False)
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
    
    ' Initialize the reference sheet
    cm.initReferences vtkReferencesInWorkbook(ActiveWorkbook)
    
    ' Add a BeforeSave event handler for the workbook
    If withBeforeSaveHandler Then vtkAddBeforeSaveHandlerInDEVWorkbook Wb:=ActiveWorkbook, projectName:=project.projectName, confName:=project.projectDEVName
    
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
'             - All configuration projectName property is the same as the project name of the configuration
'             - All configuration projectName property is the same as the title property of the configuration
'             - All configuration comment property is the same as the comment property of the configuration
'             - All configuration template path is reachable
'             - All modules listed in a configuration description are existing in the configuration
'             - All modules really present in a configuration are described in the description with non null path
'             - All modules pathes are reachable
'             - Each code module implemented in a configuration is the same as the source code module
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
            If Not (cwb(nbConf).Wb Is Nothing) Then

    ' Verify projects name
                If Not (c.projectName = cwb(nbConf).Wb.VBProject.name) Then
                    Debug.Print "For configuration " & c.name & ", the projectName property (" & c.projectName & ") is different of the project name (" & cwb(nbConf).Wb.VBProject.name & ")."
                End If

    ' Verify workbooks title
                If Not (c.projectName = cwb(nbConf).Wb.BuiltinDocumentProperties("Title").Value) Then
                    Debug.Print "For configuration " & c.name & ", the projectName property (" & c.projectName & ") is different of the workbook title (" & cwb(nbConf).Wb.BuiltinDocumentProperties("Title").Value & ")."
                End If

    ' Verify workbooks comment
                If Not (c.comment = cwb(nbConf).Wb.BuiltinDocumentProperties("Comments").Value) Then
                    Debug.Print "For configuration " & c.name & ", the comment property (" & c.comment & ") is different of the workbook comment (" & cwb(nbConf).Wb.BuiltinDocumentProperties("Comments").Value & ")."
                End If

    ' Verify workbooks template path
                If Not (fso.FileExists(cm.rootPath & "\" & c.template)) Then
                    Debug.Print "For configuration " & c.name & ", the template path (" & cm.rootPath & "\" & c.template & ") is unreachable."
                End If

               Else
                Debug.Print "Impossible to open Workbook for configuration " & c.name & " (" & s & ")."
            End If
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
    
    ' Verify that all modules pathes are reachable
    For i = 1 To nbConf
        For Each md In cwb(i).conf.modules
            s = cm.rootPath & "\" & md.path
            If Not fso.FileExists(s) Then
                Debug.Print "Module " & md.name & " path (" & md.path & " is not reachable for the configuration " & cwb(i).conf.name & "."
            End If
        Next
    Next i
    
    ' Verify that all modules content of all configuration are equal to source modules content
    ' - Create a project folder tree structure in the test folder of the cirrent project
    ' - For each configuration
    '   - Export each module in the test tree folder (to perform a comparaison on normalized export)
    '   - compare the content of the each file in the tree folder to the one in the source folder
    ' - Delete the files and folders in the test folder
    Dim testPath As String, s1 As String
    testPath = vtkPathToTestFolder(ActiveWorkbook) & "\Temporary"
    vtkCreateTreeFolder testPath
    For i = 1 To nbConf
        For Each md In cwb(i).conf.modules
            s = cm.rootPath & "\" & md.path
            s1 = testPath & "\" & md.path
            vtkExportOneModule cwb(i).Wb.VBProject, md.name, s1
            If Not compareFiles(s, s1, True) Then
                Debug.Print "Module " & md.name & " content of source path (" & md.path & " is different from module in the configuration " & cwb(i).conf.name & "."
            End If
        Next
    Next i
    vtkDeleteFolder testPath
    
    ' Verify that all references listed in a configuration description are existing in the configuration
    ' and that all references really present in a configuration are described in the description
    '   - for each configuration, get both collection
    '   - compare one list to the other whikle removing each found item
    '       - alert if an item is not found in the other list
    '   - the count of remaining items must be null at the end of the comparison
    '       - alert if not, and list the remaining items
    Dim actualList As Collection, actualRef  As vtkReference, expectedRef As vtkReference
    For i = 1 To nbConf
        Set actualList = vtkReferencesInWorkbook(cwb(i).Wb) ' Get the actual list, indexed by name
        For Each expectedRef In cwb(i).conf.references      ' the expected list is indexed by ID
           On Error Resume Next
            Set actualRef = actualList(expectedRef.name)
            If Err.Number <> 0 Then
                Debug.Print "Reference " & expectedRef.name & " is expected but not present in configuration " & cwb(i).conf.name & "."
               Else
                actualList.Remove expectedRef.name
            End If
           On Error GoTo 0
        Next expectedRef
        If actualList.Count <> 0 Then
            For Each actualRef In actualList
                Debug.Print "Reference " & actualRef.name & " is present but not expected in configuration " & cwb(i).conf.name & "."
            Next actualRef
        End If
    Next i
    
    ' Close all Worbooks opened during this verification
    For i = 1 To nbConf
        If Not cwb(i).wasOpened Then cwb(i).Wb.Close saveChanges:=False
    Next i
    Debug.Print "----------------------------------------------------"
End Sub

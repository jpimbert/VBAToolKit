Attribute VB_Name = "vtkProjectCreationUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkProjectCreationUtilities
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Utility functions used for VBAToolKit project creation
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

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeVbaUnitNamesAndPathes
' Author    : Abdelfattah Lahbib
' Date      : 09/05/2013
' Purpose   : - Initialize DEV project ConfSheet with vbaunit module names and pathes
'             - Return True if module names and paths are initialized without error
' WARNING   : This function must use an Excel ConfigurationManager (not XML)
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeVbaUnitNamesAndPathes(project As String) As Boolean
    Dim i As Integer, cm As vtkConfigurationManager, ret As Boolean, nm As Integer, nc As Integer, ext As String
    Dim moduleName As String, module As VBComponent
    
    Set cm = vtkConfigurationManagerForProject(project)
    nc = cm.getConfigurationNumber(vtkProjectForName(project).projectDEVName)
    ret = (nc > 0)
    
    For i = 1 To vtkVBAUnitModulesList.Count
        moduleName = vtkVBAUnitModulesList.Item(i)
        Set module = ThisWorkbook.VBProject.VBComponents(moduleName)
        
        nm = cm.addModule(moduleName)
        ret = ret And (nm > 0)
        
        cm.setModulePathWithNumber path:=vtkStandardPathForModule(module), numModule:=nm, numConfiguration:=nc
        
    Next i
    
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function

'---------------------------------------------------------------------------------------
' Procedure : VtkAvtivateReferences
' Author    : Jean-Pierre IMBERT
' Date      : 22/06/2014
' Purpose   : - Check that workbook is open
'             - Activate references of the configuration whose name is given as a parameter
'---------------------------------------------------------------------------------------
Public Sub VtkActivateReferences(Wb As Workbook, projectName As String, confName As String)
    Dim ref As vtkReference
    If VtkWorkbookIsOpen(Wb.name) Then
        For Each ref In vtkConfigurationManagerForProject(projectName).getConfigurationReferencesWithNumber(vtkConfigurationManagerForProject(projectName).getConfigurationNumber(confName))
            ref.addToWorkbook Wb
        Next
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkDisplayActivatedReferencesGuid
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Utility Sub for displaying GUID of activated references of current project
'---------------------------------------------------------------------------------------
'
Public Sub vtkDisplayActivatedReferencesGuid()
    Dim r As vtkReference
    For Each r In vtkReferencesInWorkbook(ActiveWorkbook)
        Debug.Print r.name, r.GUID, r.fullPath
    Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkAddBeforeSaveHandlerInDEVWorkbook
' Author    : Lucas Vitorino
' Purpose   : - Adds a Workbook_BeforeSave handler in a DEV workbook. This handler exports
'               the modified modules of the _DEV configuration associated to this workbook,
'               and exports the vtkConfigurations sheet as an XML file in the same folder.
'             - The handler will call vtkExportConfiguration on
'                 - the project of the current workbook
'                 - a project name that is the name of the worbook without "_DEV.xlsm"
'                 - a confname that is the name of the workbook without ".xslm"
'             - It works on any workbook, but shouldn't be used on something else than
'               a proper _DEV workbook of a VTKProject. Most probably, when saving the workbook,
'               an error will occur.
'---------------------------------------------------------------------------------------
'
Public Sub vtkAddBeforeSaveHandlerInDEVWorkbook(Wb As Workbook, projectName As String, confName As String)
    
    On Error GoTo vtkAddBeforeSaveHandlerInDEVWorkbook_Error
    
    ' Force the vtkReferences sheet for the BeforeSave Handler to work in eac case
    Dim c As Collection
    Set c = vtkConfigurationManagerForProject(projectName).references
    
    Dim wbVTKName As String
    wbVTKName = ThisWorkbook.VBProject.name ' Get the name of the Running project (VBAToolKit)
    
    Dim handlerString As String
    handlerString = _
    "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)" & vbNewLine & _
    "   On error goto M_Error" & vbNewLine & _
    "   " & wbVTKName & ".vtkExportConfiguration projectWithModules:=ThisWorkbook.VBProject, projectName:=" & """" & projectName & """" & _
                                                                    " , confName:=" & """" & confName & """" & _
                                                                    " , onlyModified:=True" & _
                                                                    vbNewLine & _
                                                                    vbNewLine & _
    "   " & wbVTKName & ".vtkExportConfigurationsAsXML projectName:=""" & projectName & """, filePath:=" & _
    wbVTKName & ".vtkPathOfCurrentProject(ThisWorkbook) & ""\"" & " & wbVTKName & ".vtkProjectForName(""" & projectName & """).XMLConfigurationStandardRelativePath" & vbNewLine & _
    "M_Error:" & vbNewLine & _
    "End Sub" & vbNewLine
    
    With Wb.VBProject.VBComponents("ThisWorkbook").CodeModule
        .InsertLines .CountOfLines + 1, handlerString
    End With
    
    On Error GoTo 0
    Exit Sub

vtkAddBeforeSaveHandlerInDEVWorkbook_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "vtkAddBeforeSaveHandlerInDEVWorkBook", Err.Description
    Resume Next
End Sub


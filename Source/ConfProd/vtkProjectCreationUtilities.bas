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
' Procedure : vtkDisplayActivatedReferencesGuid
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Utility Sub for displaying GUID of activated references of current project
'---------------------------------------------------------------------------------------
'
Public Sub vtkDisplayActivatedReferencesGuid()
    Dim r As VBIDE.Reference
    For Each r In ActiveWorkbook.VBProject.references
        Debug.Print r.name, r.guid
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
          
    Dim wbVTKName As String
    wbVTKName = ThisWorkbook.VBProject.name ' Get the name of the Running project (VBAToolKit)
    
    Dim handlerString As String
    handlerString = _
    "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)" & vbNewLine & _
    "   " & wbVTKName & ".vtkExportConfiguration projectWithModules:=ThisWorkbook.VBProject, projectName:=" & """" & projectName & """" & _
                                                                    " , confName:=" & """" & confName & """" & _
                                                                    " , onlyModified:=True" & _
                                                                    vbNewLine & _
                                                                    vbNewLine & _
    "   " & wbVTKName & ".vtkExportConfigurationsAsXML projectName:=""" & projectName & """, filePath:=" & _
    wbVTKName & ".vtkPathOfCurrentProject(ThisWorkbook) & ""\"" & " & wbVTKName & ".vtkProjectForName(""" & projectName & """).XMLConfigurationStandardRelativePath" & vbNewLine & _
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


'---------------------------------------------------------------------------------------
' Procedure : listOfDefaultReferences
' Author    : Lucas Vitorino
' Purpose   : Returns a collection of VBIDE.Reference objects corresponding to the default
'             references activated by default in a new VBAToolKit project.
'---------------------------------------------------------------------------------------
'
Public Function listOfDefaultReferences() As Collection
    Dim retList As Collection
    Set retList = New Collection
    
    Dim refName As Variant ' necessary evil : can't loop through a collection of String without Variant
    For Each refName In listOfDefaultReferencesNames
        retList.Add Item:=ThisWorkbook.VBProject.references(refName)
    Next
    
    Set listOfDefaultReferences = retList
End Function

'---------------------------------------------------------------------------------------
' Procedure : listOfDefaultReferencesNames
' Author    : Lucas Vitorino
' Purpose   : Returns a collection of Strings corresponding to the default
'             references activated by default in a new VBAToolKit project.
'---------------------------------------------------------------------------------------
'
Public Function listOfDefaultReferencesNames() As Collection
    Set listOfDefaultReferencesNames = New Collection
    With listOfDefaultReferencesNames
            .Add Item:="Scripting", Key:="Scripting" ' Microsoft scripting runtime
            .Add Item:="VBIDE", Key:="VBIDE" ' Microsoft visual basic for applications extensibility 5.3
            .Add Item:="Shell32", Key:="Shell32" ' Microsoft Shell Controls and Automation
            .Add Item:="MSXML2", Key:="MSXML2" ' Microsoft XML V5.0
            .Add Item:="ADODB", Key:="ADODB" ' Microsoft ActiveX Data Objects V2.6 Library
    End With
End Function


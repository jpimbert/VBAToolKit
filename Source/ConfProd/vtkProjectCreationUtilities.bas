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
        
        nm = cm.AddModule(moduleName)
        ret = ret And (nm > 0)
        
        cm.setModulePathWithNumber path:=vtkStandardPathForModule(module), numModule:=nm, numConfiguration:=nc
        
    Next i
    
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function

'---------------------------------------------------------------------------------------
' Procedure : VtkAvtivateReferences
' Author    : Abdelfattah Lahbib
' Date      : 26/04/2013
' Purpose   : - check that workbook is open and activate VBIDE and +-scripting references
'---------------------------------------------------------------------------------------
Public Sub VtkActivateReferences(wb As Workbook)
    If VtkWorkbookIsOpen(wb.name) = True Then     'if the workbook is opened
       On Error Resume Next ' if an extention is already activated, we will try to activate the next one
        wb.VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 0, 0  ' Scripting : Microsoft scripting runtime
        wb.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0  ' VBIDE : Microsoft visual basic for applications extensibility 5.3
        wb.VBProject.References.AddFromGuid "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}", 0, 0  ' Shell32 : Microsoft Shell Controls and Automation
        wb.VBProject.References.AddFromGuid "{F5078F18-C551-11D3-89B9-0000F81FE221}", 0, 0  ' MSXML2 : Microsoft XML V5.0
       On Error GoTo 0
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
    Dim r As VBIDE.Reference
    For Each r In ActiveWorkbook.VBProject.References
        Debug.Print r.name, r.GUID
    Next
End Sub

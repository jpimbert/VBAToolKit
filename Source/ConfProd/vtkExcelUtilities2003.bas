Attribute VB_Name = "vtkExcelUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkExcelUtilities
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Utilities for Excel file management
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
' Procedure : vtkCreateExcelWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel file creation
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbook() As Workbook
    Dim Wb As Workbook
    Set Wb = Workbooks.Add(xlWBATWorksheet)
    Set vtkCreateExcelWorkbook = Wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelWorkbookForTestWithProjectName
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel file creation for Test and VBA project name initialization
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbookForTestWithProjectName(projectName As String) As Workbook
    Dim Wb As Workbook
    Set Wb = vtkCreateExcelWorkbookWithPathAndName(vtkPathToTestFolder, vtkProjectForName(projectName).workbookDEVName)
    Wb.VBProject.name = projectName
    Set vtkCreateExcelWorkbookForTestWithProjectName = Wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelWorkbookWithPathAndName
' Author    : Jean-Pierre Imbert
' Date      : 08/06/2013
' Purpose   : Create a New Excel File and save it on the given path with the given name
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbookWithPathAndName(path As String, name As String) As Workbook
    Dim Wb As Workbook
    Set Wb = vtkCreateExcelWorkbook
    Wb.SaveAs fileName:=path & "\" & name, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Set vtkCreateExcelWorkbookWithPathAndName = Wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCloseAndKillWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 08/06/2013
' Purpose   : Close the given workbook then kill the Excel File
'---------------------------------------------------------------------------------------
'
Public Sub vtkCloseAndKillWorkbook(Wb As Workbook)
    Dim fullPath As String
    fullPath = Wb.FullName
    Wb.Close saveChanges:=False
    Kill PathName:=fullPath
End Sub

'---------------------------------------------------------------------------------------
' Procedure : VtkWorkbookIsOpen
' Author    : Abdelfattah Lahbib
' Date      : 24/04/2013
' Purpose   : return true if the workbook is already opened
'---------------------------------------------------------------------------------------
'
Public Function VtkWorkbookIsOpen(workbookName As String) As Boolean
    On Error Resume Next
    Workbooks(workbookName).Activate 'if we have a problem to activate workbook = the workbook is closed , if we can activate it without problem = the workbook is open
    VtkWorkbookIsOpen = (Err = 0)
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDefaultFileFormat
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : return the default FileFormat of an Excel file given its full path
'             - xlOpenXMLWorkbookMacroEnabled for .xlsm
'             - xlOpenXMLAddIn for .xlam
'             - 0 if the extension is unknown
'---------------------------------------------------------------------------------------
'
Public Function vtkDefaultFileFormat(filePath As String) As XlFileFormat
    Select Case vtkGetFileExtension(filePath)
        Case "xla"
            vtkDefaultFileFormat = xlAddIn
        Case "xls"
            vtkDefaultFileFormat = xlWorkbookNormal
        Case Else
            vtkDefaultFileFormat = 0
        End Select
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkDefaultExcelExtension as String
' Author    : Jean-Pierre Imbert
' Date      : 06/09/2014
' Purpose   : return the default extension for an Excel file
'             - .xls for Excel until 2003 (11.0)
'             - .xlsm for Excel since 2007 (12.0)
'---------------------------------------------------------------------------------------
'
Public Function vtkDefaultExcelExtension() As String
    If Val(Application.Version) >= 12 Then
        vtkDefaultExcelExtension = ".xlsm"
    Else
        vtkDefaultExcelExtension = ".xls"
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkDefaultIsAddIn as Boolean
' Author    : Jean-Pierre Imbert
' Date      : 11/11/2013
' Purpose   : return True if the default File Format is an Excel Add-In
'             - currenttly is add-in if extension is .xla or .xlam
'---------------------------------------------------------------------------------------
'
Public Function vtkDefaultIsAddIn(filePath As String) As Boolean
    vtkDefaultIsAddIn = (vtkGetFileExtension(filePath) = "xlam") Or (vtkGetFileExtension(filePath) = "xla")
End Function

'---------------------------------------------------------------------------------------
' Function  : vtkReferencesInWorkbook as Collection of vtkReferences
' Author    : Jean-Pierre Imbert
' Date      : 23/06/2014
' Purpose   : return a collection of references of the workbook given as a parameter
'             - the ID property of each vtkReference instance of this collection is not initialized
'             - the collection is indexed by the name of the references
' NOTE      : this function is not tested
'---------------------------------------------------------------------------------------
'
Public Function vtkReferencesInWorkbook(Wb As Workbook) As Collection
    Dim c As New Collection, ref As vtkReference, r As VBIDE.Reference
    For Each r In Wb.VBProject.references
        Set ref = New vtkReference
        ref.initWithVBAReference r
        c.Add Item:=ref, Key:=ref.name
    Next
    Set vtkReferencesInWorkbook = c
End Function

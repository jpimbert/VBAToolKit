Attribute VB_Name = "vtkExcelUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel file creation
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbook() As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set vtkCreateExcelWorkbook = wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelWorkbookForTestWithProjectName
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel file creation for Test and VBA project name initialization
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbookForTestWithProjectName(projectName As String) As Workbook
    Dim wb As Workbook
    Set wb = vtkCreateExcelWorkbookWithPathAndName(vtkPathToTestFolder, vtkProjectForName(projectName).workbookDEVName)
    wb.VBProject.name = projectName
    Set vtkCreateExcelWorkbookForTestWithProjectName = wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelWorkbookWithPathAndName
' Author    : Jean-Pierre Imbert
' Date      : 08/06/2013
' Purpose   : Create a New Excel File and save it on the given path with the given name
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbookWithPathAndName(path As String, name As String) As Workbook
    Dim wb As Workbook
    Set wb = vtkCreateExcelWorkbook
    wb.SaveAs Filename:=path & "\" & name, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Set vtkCreateExcelWorkbookWithPathAndName = wb
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCloseAndKillWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 08/06/2013
' Purpose   : Close the given workbook then kill the Excel File
'---------------------------------------------------------------------------------------
'
Public Sub vtkCloseAndKillWorkbook(wb As Workbook)
    Dim fullPath As String
    fullPath = wb.FullName
    wb.Close SaveChanges:=False
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
    VtkWorkbookIsOpen = (err = 0)
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
        Case "xlsx"
            vtkDefaultFileFormat = xlOpenXMLWorkbook
        Case "xltx"
            vtkDefaultFileFormat = xlOpenXMLTemplate
        Case "xltm"
            vtkDefaultFileFormat = xlOpenXMLTemplateMacroEnabled
        Case "xlsm"
            vtkDefaultFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xlam"
            vtkDefaultFileFormat = xlOpenXMLAddIn
        Case "xla"
            vtkDefaultFileFormat = xlAddIn
        Case Else
            vtkDefaultFileFormat = 0
        End Select
End Function

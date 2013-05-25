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
' Procedure : vtkCreateExcelWorkbookWithProjectName
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel file creation and VBA project name initialization
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelWorkbookWithProjectName(projectName As String) As Workbook
    Dim wb As Workbook
    Set wb = vtkCreateExcelWorkbook
    wb.VBProject.name = projectName
    Set vtkCreateExcelWorkbook = wb
End Function


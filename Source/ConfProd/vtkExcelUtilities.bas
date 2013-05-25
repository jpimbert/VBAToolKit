Attribute VB_Name = "vtkExcelUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : vtkCreateExcelProjectNamed
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Utility function for Excel project creation with a given project name
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateExcelProjectNamed(projectName As String) As Workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.VBProject.name = projectName
    Set vtkCreateExcelProjectNamed = wb
End Function



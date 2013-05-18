Attribute VB_Name = "vtkImportExportUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Procedure : vtkConfSheet
' Author    : user
' Date      : 14/05/2013
' Purpose   : - create new sheet (if it not exist) that will contain table of parameters
'---------------------------------------------------------------------------------------
'
Public Function vtkConfSheet() As String
Dim sheetname
sheetname = "configurations"
On Error Resume Next
 Worksheets(sheetname).Select
 If Err <> 0 Then
 Worksheets.Add.name = sheetname
 End If
vtkConfSheet = sheetname
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleNameRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of modules
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleNameRange() As String
vtkModuleNameRange = "A"
ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkModuleNameRange & vtkFirstLine - 2) = "Module Name"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDevRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of path of developemnt configuration
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDevRange() As String
vtkModuleDevRange = "B"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkModuleDeliveryRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains list of path of devivery configuration
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkModuleDeliveryRange() As String
vtkModuleDeliveryRange = "C"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkInformationRange
' Author    : user
' Date      : 13/05/2013
' Purpose   : - return range name that contains modules information
'             - write range name
'---------------------------------------------------------------------------------------
'
Public Function vtkInformationRange() As String
vtkInformationRange = "D"
ActiveWorkbook.Sheets(vtkConfSheet).Range(vtkInformationRange & vtkFirstLine - 3) = "File Informations"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkFirstLine
' Author    : user
' Date      : 13/05/2013
' Purpose   : - define the start line
'---------------------------------------------------------------------------------------
'
Public Function vtkFirstLine() As Integer
vtkFirstLine = 4
End Function

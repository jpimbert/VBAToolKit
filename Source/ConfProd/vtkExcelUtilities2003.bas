Attribute VB_Name = "vtkExcelUtilitiesSpecific"
'---------------------------------------------------------------------------------------
' Module    : vtkExcelUtilitiesSpecific
' Author    : Jean-Pierre Imbert
' Date      : 16/01/2015
' Purpose   : Specific Utilities (2003-2007) for Excel file management
'
' Copyright 2015 Skwal-Soft (http://skwalsoft.com)
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
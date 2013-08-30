Attribute VB_Name = "vtkTestUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkTestUtilities
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Some utilities to facilitate test writing
'             - vtkTestPath, gives the path of the current project
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


'---------------------------------------------------------------------------------------
' Procedure : vtkTestPath
' Author    : Jean-Pierre Imbert
' Date      : 07/05/2013
' Purpose   : Return the path of the Test Folder of the current project  '.\Tests
'---------------------------------------------------------------------------------------
'
Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder
End Function

'---------------------------------------------------------------------------------------
' Procedure : getTestFileFromTemplate
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Copy a File from the Template folder to the Test folder and optionaly open it
' Parameters
'           - fileName as string, file to get from the Template folder
'           - Optional destinationName as string, name of file to create in the Test folder (same as fileName by default)
'           - Optional openExcel as Boolean, if True open the file as Excel workbook, false by default
' Return    : The opened Excel workbook or Nothing if no open file or error during opening
'
' Note      : In case of Err 1004, 5 retries are attempted before return Nothing
'             The Err 1004 can be raised if the file copy is not completely performed before opening
'---------------------------------------------------------------------------------------
'
Public Function getTestFileFromTemplate(fileName As String, Optional destinationName As String = "", Optional openExcel As Boolean = False) As Workbook
    Dim source As String, destination As String, errCount As Integer
    
    ' Copy file
    source = vtkPathToTemplateFolder & "\" & fileName
    If destinationName Like "" Then
        destination = vtkTestPath & "\" & fileName
       Else
        destination = vtkTestPath & "\" & destinationName
    End If
    FileCopy source:=source, destination:=destination
    
    ' Open Excel file if required
    Set getTestFileFromTemplate = Nothing
    If openExcel Then
        errCount = 0
       On Error GoTo M_Error
        Set getTestFileFromTemplate = Workbooks.Open(destination)
       On Error GoTo 0
    End If
    Exit Function

M_Error:
    errCount = errCount + 1
    If err.Number = 1004 And errCount < 5 Then Resume    ' It's possible that the file is not ready, just after copy : in this case retry
    Set getTestFileFromTemplate = Nothing
    err.Raise Number:=err.Number, source:=err.source, Description:=err.Description
End Function

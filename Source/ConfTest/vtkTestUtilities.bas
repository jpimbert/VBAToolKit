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

Private pWorkBook As Workbook

'---------------------------------------------------------------------------------------
' Procedure : prepare
' Author    : Jean-Pierre IMBERT
' Date      : 31/08/2013
' Purpose   : Prepare the module before use in test
'---------------------------------------------------------------------------------------
'
Public Sub prepare(Wb As Workbook)
    Set pWorkBook = Wb    ' VBAToolKit works on Active Workbook by default
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkTestPath
' Author    : Jean-Pierre Imbert
' Date      : 07/05/2013
' Purpose   : Return the path of the Test Folder of the current project  '.\Tests
'---------------------------------------------------------------------------------------
'
Public Function vtkTestPath() As String
    vtkTestPath = vtkPathToTestFolder(pWorkBook)
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
' Error raised :
'           - VTK_FILE_NOT_FOUND, in case of file name not found in template folder
'           - VTK_UNEXPECTED_ERROR, all other case
'---------------------------------------------------------------------------------------
'
Public Function getTestFileFromTemplate(fileName As String, Optional destinationName As String = "", Optional openExcel As Boolean = False) As Workbook
    Dim Source As String, destination As String, errCount As Integer
    
   On Error GoTo M_Error
   
    ' Copy file
    Source = vtkPathToTemplateFolder(pWorkBook) & "\" & fileName
    If destinationName Like "" Then
        destination = vtkTestPath & "\" & fileName
       Else
        destination = vtkTestPath & "\" & destinationName
    End If
    
    FileCopy Source:=source, destination:=destination
    
    ' Open Excel file if required
    Set getTestFileFromTemplate = Nothing
    If openExcel Then
        errCount = 0
        Set getTestFileFromTemplate = Workbooks.Open(destination)
    End If
    
   On Error GoTo 0
    Exit Function

M_Error:
    errCount = errCount + 1
    If Err.Number = 1004 And errCount < 5 Then Resume    ' It's possible that the file is not ready, just after copy : in this case retry
    Set getTestFileFromTemplate = Nothing
    Select Case Err.Number
        Case 53
            Err.Raise Number:=VTK_FILE_NOT_FOUND, Source:="getTestFileFromTemplate", Description:="File not found : " & Source
        Case 75
            Err.Raise Number:=VTK_DOESNT_COPY_FOLDER, Source:="getTestFileFromTemplate", Description:="A file can't be copied : " & Source
        Case Else
            Err.Raise VTK_UNEXPECTED_ERROR, "getTestFileFromTemplate", "(" & Err.Number & ") " & Err.Description
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure  : vtkGetTestFolderFromTemplate
' Author     : Champonnois
' Date       : 23/09/2013
' Purpose    : Copy a folder from the Template folder to the Test folder
' Parameters :
'           - fileName as string, folder to get from the Template folder
'           - Optional destinationName as string, name of folder to create in the Test folder (same as folderName by default)
'---------------------------------------------------------------------------------------
Public Function getTestFolderFromTemplate(folderName As String, Optional destinationName As String = "")

    Dim Source As String, destination As String, errCount As Integer, fso As FileSystemObject

    On Error GoTo getTestFolderFromTemplate_Error
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Copy folder
    Source = vtkPathToTemplateFolder(pWorkBook) & "\" & folderName
    
    If destinationName Like "" Then
        destination = vtkTestPath & "\" & folderName
       Else
        destination = vtkTestPath & "\" & destinationName
    End If
    
    fso.CopyFolder Source:=source, destination:=destination, OverWriteFiles:=True
    
    On Error GoTo 0
    Exit Function

getTestFolderFromTemplate_Error:
    Select Case Err.Number
        Case 76
            Err.Raise Number:=VTK_FOLDER_NOT_FOUND, Source:="getTestFolderFromTemplate", Description:="Folder not found : " & Source
        Case Else
            Err.Raise VTK_UNEXPECTED_ERROR, "getTestFolderFromTemplate", "(" & Err.Number & ") " & Err.Description
    End Select
    Resume Next
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResetTestFolder
' Author    : Champonnois
' Date      : 25/09/2013
' Purpose   : Remove the contents of the folder test
'
' Raise error :
'           - VTK_FILE_OPEN_OR_LOCKED, the folder can't be clean up
'           - VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Sub resetTestFolder()
    Dim fso As New FileSystemObject
   On Error GoTo resetTestFolder_Error

    fso.DeleteFolder VBAToolKit.vtkTestPath & "\*", Force:=True
    fso.DeleteFile VBAToolKit.vtkTestPath & "\*.*", Force:=True

   On Error GoTo 0
   Exit Sub

resetTestFolder_Error:
    Select Case Err.Number
        Case 70
            Err.Raise Number:=VTK_FILE_OPEN_OR_LOCKED, Source:="resetTestFolder", Description:=Err.Description
        Case Else
            Err.Raise VTK_UNEXPECTED_ERROR, "resetTestFolder", "(" & Err.Number & ") " & Err.Description
    End Select
End Sub

'---------------------------------------------------------------------------------------
' Procedure : insertDummyProcedureInCodeModule
' Author    : Lucas Vitorino
' Purpose   : - Insert a dummy procedure at the end of a VBIDE.CodeModule object
'             - The optional argument allows adding a number to the name of the procedure
'               so as to avoid same-name procedures in the same module.
'---------------------------------------------------------------------------------------
'
Public Sub insertDummyProcedureInCodeModule(codemo As VBIDE.CodeModule, Optional dummyInt As Integer = 0)
    Dim dummyProcedure As String
    
    On Error GoTo insertDummyProcedureInCodeModule_Error

    dummyProcedure = _
    "Public Sub dummyProcedure" & dummyInt & "()" & vbNewLine & _
    "End Sub" & vbNewLine
    
    With codemo
        .InsertLines .CountOfLines + 1, dummyProcedure
    End With

    On Error GoTo 0
    Exit Sub

insertDummyProcedureInCodeModule_Error:
    Err.Raise VTK_UNEXPECTED_ERROR, "sub insertDummyProcedureInCodeModule of module vtkTestUtilities", Err.Description
    Resume Next

End Sub

'---------------------------------------------------------------------------------------
' Procedure : compareFiles
' Author    : http://www.freevbcode.com
' Date      : 14/11/2013
' Purpose   : Check to see if two files are identical
'             File1 and File2 = FullPaths of files to compare
'             StringentCheck (optional): If false (default),
'             will only compare file lengths.  If true, a
'             byte by byte comparison is conducted if file lengths are
'             equal
'---------------------------------------------------------------------------------------
'
Public Function compareFiles(ByVal File1 As String, _
  ByVal File2 As String, Optional StringentCheck As _
  Boolean = False) As Boolean

On Error GoTo ErrorHandler

If Dir(File1) = "" Then Exit Function
If Dir(File2) = "" Then Exit Function

Dim lLen1 As Long, lLen2 As Long
Dim iFileNum1 As Integer
Dim iFileNum2 As Integer
Dim bytArr1() As Byte, bytArr2() As Byte
Dim lCtr As Long, lStart As Long
Dim bAns As Boolean

lLen1 = FileLen(File1)
lLen2 = FileLen(File2)
If lLen1 <> lLen2 Then
    compareFiles = False
    Exit Function
ElseIf StringentCheck = False Then
        compareFiles = True
        Exit Function
Else
    iFileNum1 = FreeFile
    Open File1 For Binary Access Read As #iFileNum1
    iFileNum2 = FreeFile
    Open File2 For Binary Access Read As #iFileNum2

    'put contents of both into byte Array
    bytArr1() = InputB(LOF(iFileNum1), #iFileNum1)
    bytArr2() = InputB(LOF(iFileNum2), #iFileNum2)
    lLen1 = UBound(bytArr1)
    lStart = LBound(bytArr1)
    
    bAns = True
    For lCtr = lStart To lLen1
        If bytArr1(lCtr) <> bytArr2(lCtr) Then
            bAns = False
            Exit For
        End If
            
    Next
    compareFiles = bAns
       
End If
 
ErrorHandler:
If iFileNum1 > 0 Then Close #iFileNum1
If iFileNum2 > 0 Then Close #iFileNum2
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCompareConfManagers
' Author    : Jean-Pierre Imbert
' Date      : 15/07/2014
' Purpose   : Compare Conf managers and use mAssert to report differences
'---------------------------------------------------------------------------------------
'
Public Sub vtkCompareConfManagers(ByVal mAssert As IAssert, ByVal expectedConf As vtkConfigurationManager, ByVal actualConf As vtkConfigurationManager)
    Dim refX As vtkReference, refE As vtkReference, i As Integer, j As Integer
    ' Check configuration manager parameters
    mAssert.Equals actualConf.projectName, expectedConf.projectName, "Project name of conf manager"
    mAssert.Equals actualConf.rootPath, expectedConf.rootPath, "Root Path of conf manager"
    ' Check counts
    mAssert.Equals actualConf.configurationCount, expectedConf.configurationCount, "Configuration count of conf manager"
    mAssert.Equals actualConf.moduleCount, expectedConf.moduleCount, "Module count of conf manager"
    mAssert.Equals actualConf.references.count, expectedConf.references.count, "Reference count of conf manager"
    ' Check configurations name and parameters
    For i = 1 To expectedConf.configurationCount
        mAssert.Equals actualConf.configuration(i), expectedConf.configuration(i), "Name of configuration number " & i
        mAssert.Equals actualConf.getConfigurationPathWithNumber(i), expectedConf.getConfigurationPathWithNumber(i), "Configuration Path of configuration number " & i
        mAssert.Equals actualConf.getConfigurationProjectNameWithNumber(i), expectedConf.getConfigurationProjectNameWithNumber(i), "Project name of configuration number " & i
        mAssert.Equals actualConf.getConfigurationTemplateWithNumber(i), expectedConf.getConfigurationTemplateWithNumber(i), "Template path of configuration number " & i
        mAssert.Equals actualConf.getConfigurationCommentWithNumber(i), expectedConf.getConfigurationCommentWithNumber(i), "Comment of configuration number " & i
    Next i
    ' Check modules name
    For i = 1 To expectedConf.moduleCount
        mAssert.Equals actualConf.module(i), expectedConf.module(i), "Name of module number " & i
    Next i
    ' Check references name and parameters
    For i = 1 To expectedConf.references.Count
        Set refX = expectedConf.references(i)
        Set refE = actualConf.references(i)
        mAssert.Equals refE.name, refX.name, "Name of reference number " & i
        mAssert.Equals refE.relPath, refX.relPath, "Relative path of reference number " & i
        mAssert.Equals refE.GUID, refX.GUID, "GUID of reference number " & i
    Next i
    ' Check module pathes
    For i = 1 To expectedConf.configurationCount
        For j = 1 To expectedConf.moduleCount
            mAssert.Equals actualConf.getModulePathWithNumber(j, i), expectedConf.getModulePathWithNumber(j, i), "Module path for module " & j & " and configuration " & i
        Next j
    Next i
    ' Check reference uses
    For i = 1 To expectedConf.configurationCount
        mAssert.Equals actualConf.getConfigurationReferencesWithNumber(i).count, expectedConf.getConfigurationReferencesWithNumber(i).count, "Reference count for configuration number " & i
        For j = 1 To expectedConf.getConfigurationReferencesWithNumber(i).Count
            Set refX = expectedConf.getConfigurationReferencesWithNumber(i)(j)
           On Error Resume Next
            Set refE = actualConf.getConfigurationReferencesWithNumber(i)(j)
           On Error GoTo 0
            mAssert.Equals refE.id, refX.id, "Reference Id for used reference number " & j & " for configuration number " & i
        Next j
    Next i
End Sub



Attribute VB_Name = "vtkFileSystemUtilities"
'---------------------------------------------------------------------------------------
' Module    : vtkFileSystemUtilities
' Author    : Lucas Vitorino
' Purpose   : Provide some utilities for interacting with files and folders.
'               - creation
'               - existence
'               - reading
'               - deletion...
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : vtkTextFileReader
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : Returns the content of a text file
' Notes     : Notably used to read Git log files.
'---------------------------------------------------------------------------------------
'
Public Function vtkTextFileReader(fullFilePath As String) As String

    Dim Textfile As Variant
    Dim strresult As String
    Dim fso As New FileSystemObject

On Error GoTo vtkTextFileReader_Error

    Set Textfile = fso.OpenTextFile(fullFilePath, ForReading)
    'while not end of file
    Do Until Textfile.AtEndOfStream
    'read line per line
        strresult = strresult & Chr(10) & Textfile.ReadLine
    Loop
    'return file text
    vtkTextFileReader = strresult

   On Error GoTo 0
   Exit Function

vtkTextFileReader_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VtkTextFileReader of Module vtkGitFunctions"
    vtkTextFileReader = Err.Number
    
End Function


'---------------------------------------------------------------------------------------
' Function  : vtkCreateTreeFolder
' Author    : Jean-Pierre Imbert
' Date      : 06/08/2013
' Purpose   : Create a project folder breakdown into the folder given as parameter
'             This procedure is isolated to be easier to test
' Return    : Long error number
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateTreeFolder(rootPath As String)
   On Error GoTo vtkCreateTreeFolder_Error
    ' Create main folder
    MkDir rootPath
    ' Create Delivery folder
    MkDir rootPath & "\" & "Delivery"
    ' Create Project folder
    MkDir rootPath & "\" & "Project"
    ' Create Tests folder
    MkDir rootPath & "\" & "Tests"
    ' Create GitLog Folder
    MkDir rootPath & "\" & "GitLog"
    ' Create Source folder
    MkDir rootPath & "\" & "Source"
    ' Create ConfProd folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfProd"
    ' Create ConfTest folder
    MkDir rootPath & "\" & "Source" & "\" & "ConfTest"
    ' Create VbaUnit folder
    MkDir rootPath & "\" & "Source" & "\" & "VbaUnit"

   On Error GoTo 0
   vtkCreateTreeFolder = 0
   Exit Function
vtkCreateTreeFolder_Error:
    vtkCreateTreeFolder = Err.Number
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteTreeFolder
' Author    : Jean-Pierre Imbert
' Date      : 06/08/2013
' Purpose   : Delete a project folder breakdown given as parameter
'             This procedure is for test purpose
'---------------------------------------------------------------------------------------
'
Public Sub vtkDeleteTreeFolder(rootPath As String)
    Dir (rootPath)                  ' Make sure to be out of the folder to clean it without Err
    On Error Resume Next
    Kill rootPath & "\Source\ConfProd\*"
    RmDir rootPath & "\Source\ConfProd"
    Kill rootPath & "\Source\ConfTest\*"
    RmDir rootPath & "\Source\ConfTest"
    Kill rootPath & "\Source\VbaUnit\*"
    RmDir rootPath & "\Source\VbaUnit"
    Kill rootPath & "\GitLog\*"
    RmDir rootPath & "\GitLog"
    Kill rootPath & "\Tests\*"
    RmDir rootPath & "\Tests"
    Kill rootPath & "\Source\*"
    RmDir rootPath & "\Source"
    Kill rootPath & "\Delivery\*"
    RmDir rootPath & "\Delivery"
    Kill rootPath & "\Project\*"
    RmDir rootPath & "\Project"
    Kill rootPath & "\*"
    RmDir rootPath
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkCleanFolder
' Author    : Lucas Vitorino
' Purpose   : Recursively delete all the content of a folder, leaving it empty.
' Notes     : Returns
'               - VTK_RETVAL_OK if successful
'               - 76 if wrong path or parameter is not a folder
'               - VTK_RETVAL_UNEXPECTED_ERR if other error
'---------------------------------------------------------------------------------------
'
Public Function vtkCleanFolder(folderPath As String) As Integer
    
    On Error GoTo vtkCleanFolder_Error
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    Dim file As Scripting.file

    ' Will raise an error if folderPath does not correspond to a valid folder
    Set sourceFolder = fso.GetFolder(folderPath)

    ' Erase the files in the folder, even the hidden ones
    For Each file In sourceFolder.Files
        fso.DeleteFile file
    Next file
    
    ' Call the function on all the SubFolders
    For Each subFolder In sourceFolder.SubFolders
        vtkCleanFolder (subFolder.path)
        fso.DeleteFolder subFolder
    Next subFolder
    
    On Error GoTo 0
    vtkCleanFolder = VTK_OK
    Exit Function
    
vtkCleanFolder_Error:
    If Err.Number = 76 Then
        vtkCleanFolder = Err.Number
    Else
        vtkCleanFolder = VTK_UNEXPECTED_ERROR
    End If
    Exit Function
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteFolder
' Author    : Lucas Vitorino
' Purpose   : Delete a folder and its content.
' Notes     : Returns
'               - VTK_RETVAL_OK if successful
'               - 76 if wrong path or parameter is not a folder
'               - VTK_RETVAL_UNEXPECTED_ERR if other error
'---------------------------------------------------------------------------------------
'
Public Function vtkDeleteFolder(folderPath As String)
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder

    On Error GoTo vtkDeleteFolder_Error

    'Will raise an error if the folder doesn't exist
    Set sourceFolder = fso.GetFolder(folderPath)
    
    vtkCleanFolder (folderPath)
    fso.DeleteFolder (sourceFolder.path)

    On Error GoTo 0
    vtkDeleteFolder = VTK_OK
    Exit Function

vtkDeleteFolder_Error:
    If Err.Number = 76 Then
        vtkDeleteFolder = Err.Number
    Else
        vtkDeleteFolder = VTK_UNEXPECTED_ERROR
    End If
    Exit Function

End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkDoesFolderExist
' Author    : Lucas Vitorino
' Purpose   : Checks if a folder exists.
' Returns   : Boolean. True if the folder exists, hidden or not, False in other cases.
'---------------------------------------------------------------------------------------
'
Public Function vtkDoesFolderExist(folderPath As String) As Integer

    On Error GoTo vtkDoesFolderExist_Error
    
    'Dir(etc,vbDirectory) returns True even if the specified thing is not a directory
    Dim fso As New FileSystemObject
    'Will raise an error 76 if wrong path or not a folder
    fso.GetFolder (folderPath)
    
    On Error GoTo 0
    vtkDoesFolderExist = True
    Exit Function

vtkDoesFolderExist_Error:
    vtkDoesFolderExist = False
    Exit Function

End Function


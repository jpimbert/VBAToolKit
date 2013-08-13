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
' Procedure : vtkCreateFileInDirectory
' Author    : Abdelfattah Lahbib
' Date      : 30/04/2013
' Purpose   : Create a file named $fileName in the directory $directory
' Notes     : Notably used for creating Git log files
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateFileInDirectory(fileName As String, directory As String) As String

    Dim fso As New FileSystemObject
    Dim fullFilePath As String
      
On Error GoTo vtkCreateFileInDirectory_Error
    
    fullFilePath = directory & "\" & fileName
      
    ' If the file doesn't exist, we create it
    If fso.FileExists(fullFilePath) = False Then
            fso.CreateTextFile (fullFilePath)
    End If
      
    'return full created file path
    vtkCreateFileInDirectory = fullFilePath

    On Error GoTo 0
    Exit Function

vtkCreateFileInDirectory_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure vtkCreateFileInDirectory of Module vtkGitFunctions"
    vtkCreateFileInDirectory = err.Number

End Function

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
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure VtkTextFileReader of Module vtkGitFunctions"
    vtkTextFileReader = err.Number
    
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
    vtkCreateTreeFolder = err.Number
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
' Procedure : vtkCleanFolder_Dir
' Author    : Lucas Vitorino
' Purpose   : Recursively delete all the content of a folder, leaving it empty.
' Notes     : - uses only the Dir function rather than FileSystemObject
'             - WIP. Doesn't currently work
'---------------------------------------------------------------------------------------
'
Public Function vtkCleanFolder_Dir(folderPath As String) As Long
    
    On Error GoTo vtkCleanFolder_Error
    
    Do Until vtkIsFolderEmpty(folderPath) = True
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    Debug.Print "IN FOLDER : " & folderPath
    
    Dim subFolder As String
    
    ' Erase the files in the folder
    Kill folderPath & "\*"
    
    subFolder = Dir(folderPath, vbDirectory)
    
        ' Until there is not subfolder left
        Do Until subFolder = ""
            If (GetAttr(folderPath & subFolder) And vbDirectory) Then
                If subFolder <> "." And subFolder <> ".." Then
                    vtkCleanFolder (folderPath & subFolder)
                    Debug.Print "Erasing folder : " & folderPath & subFolder
                    RmDir folderPath & subFolder
                End If
            End If
            Debug.Print "DEBUG <<"
            Debug.Print "Current folder = " & folderPath
            Debug.Print "Current subfolder = " & subFolder
            subFolder = Dir()
            Debug.Print ">> DEBUG "
        Loop
    
    Loop
    Debug.Print "Folder " & folderPath & " empty"
    
    On Error GoTo 0
    vtkCleanFolder = VTK_RETVAL_OK
    Exit Function
    
vtkCleanFolder_Dir_Error:
    ' Kill sourceFolder.path & "\*" will throw an error 53 if the folder is empty.
    If err.Number = 53 Then
        Resume Next
    Else
        vtkCleanFolder = VTK_RETVAL_UNEXPECTED_ERROR
        Debug.Print "ERROR " & err.Number & " : " & err.Description
        Exit Function
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkCleanFolder
' Author    : Lucas Vitorino
' Purpose   : Recursively delete all the content of a folder, leaving it empty.
' Notes     : - uses Scripting.FileSystemObject
'---------------------------------------------------------------------------------------
'
Public Function vtkCleanFolder(folderPath As String) As Integer
    
    On Error GoTo vtkCleanFolder_Error
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder
    Dim subFolder As Scripting.Folder
    
    Set sourceFolder = fso.GetFolder(folderPath)
    
    ' Erase the files in the folder
    Kill sourceFolder.path & "\*"
    
    ' Call the function on all the SubFolders
    For Each subFolder In sourceFolder.SubFolders
        vtkCleanFolder (subFolder.path)
        RmDir subFolder.path
    Next subFolder
    
    On Error GoTo 0
    vtkCleanFolder = VTK_RETVAL_OK
    Exit Function
    
vtkCleanFolder_Error:
    ' Kill sourceFolder.path & "\*" will throw an error 53 if the folder is empty.
    If err.Number = 53 Then
        Resume Next
    Else
        vtkCleanFolder = VTK_RETVAL_UNEXPECTED_ERROR
        Exit Function
    End If
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteFolder
' Author    : Lucas Vitorino
' Purpose   : Delete a folder and its content.
'---------------------------------------------------------------------------------------
'
Public Sub vtkDeleteFolder(folderPath As String)
    vtkCleanFolder folderPath
    RmDir folderPath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkIsFolderEmpty
' Author    : Jean-Pierre Imbert
' Purpose   : Checks if a folder is empty (no subfolders, no files)
' Return    : Boolean
'---------------------------------------------------------------------------------------
'
Public Function vtkIsFolderEmpty(folderPath As String)
    
    Dim fso As New Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder
    Set sourceFolder = fso.GetFolder(folderPath)
    
    vtkIsFolderEmpty = ((sourceFolder.SubFolders.Count = 0) And Dir(folderPath & "\*.*") = "")
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkDoesFolderExist
' Author    : Lucas Vitorino
' Purpose   : Checks if a folder exists.
'---------------------------------------------------------------------------------------
'
Public Function vtkDoesFolderExist(folderPath As String) As Boolean
    If Dir(folderPath, vbDirectory) = "" Then
        vtkDoesFolderExist = False
    Else
        vtkDoesFolderExist = True
    End If
End Function


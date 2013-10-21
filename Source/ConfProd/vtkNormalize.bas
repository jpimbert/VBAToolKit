Attribute VB_Name = "vtkNormalize"
'---------------------------------------------------------------------------------------
' Module    : vtkNormalize
' Author    : Lucas Vitorino
' Purpose   : This module contains the functions called for the normalization of source code.
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

Private properlyCasedIdentifiersArray() As String

'---------------------------------------------------------------------------------------
' Procedure : initializeList
' Author    : Lucas Vitorino
' Purpose   : This Sub initializes the list of properly cased identifiers used to normalize the source code.
'---------------------------------------------------------------------------------------
'
Private Sub initializeList()
    
    Dim properlyCasedIdentifiersString As String
    
    properlyCasedIdentifiersString = _
    "Dim" & "," & _
    "Wb" & "," & _
    "Err" & "," & _
    "Number" & "," & _
    "Description" & "," & _
    "Source" & "," & _
    "Count" & "," & _
    "File" & "," & _
    "Folder" & "," & _
    "Scripting" & "," & _
    "Boolean" & "," & _
    "String" & "," & _
    "Integer" & "," & _
    "addModule" & "," & _
    "returnValue" & "," & _
    "retVal" & "," & _
    "fileName" & "," & _
    "saveChanges" & "," & _
    "FileExists" & "," & _
    "m" & "," & _
    "c" & "," & _
    "Tests" & "," & _
    ""
    
    properlyCasedIdentifiersArray = Split(properlyCasedIdentifiersString, ",")
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkListOfProperlyCasedIdentifiers
' Author    : Lucas Vitorino
' Purpose   : This functions returns the array containing the properly cased Strings.
'---------------------------------------------------------------------------------------
'
Public Function vtkListOfProperlyCasedIdentifiers() As String()
    
    If Len(Join(properlyCasedIdentifiersArray, "")) = 0 Then ' if the array has not been initialized
        initializeList
    End If
    
    vtkListOfProperlyCasedIdentifiers = properlyCasedIdentifiersArray
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkNormalizeToken
' Author    : Lucas Vitorino
' Purpose   : - Perform a case-insensitive search of a given token in a given array of Strings,
'               and return this token with the casing used in the array.
'             - If the token is not found in the array of if the array is empty, the token is
'               returned unchanged.
' Returns   : A String corresponding to the normalized token.
' Errors    : VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Function vtkNormalizeToken(token As String, listOfTokens() As String) As String

    On Error GoTo vtkNormalizeToken_Error
    
    ' If the list is not empty
    If Len(Join(listOfTokens, "")) <> 0 Then
        ' Perform the search
        Dim i As Integer
        For i = LBound(listOfTokens) To UBound(listOfTokens)
            If StrComp(UCase(token), UCase(listOfTokens(i))) = 0 Then
                vtkNormalizeToken = listOfTokens(i)
                Exit Function
            End If
        Next
    End If
    
    ' If the token has not been found or the list is empty, return the token unchanged
    On Error GoTo 0
    vtkNormalizeToken = token
    Exit Function

vtkNormalizeToken_Error:
    Err.Source = "function vtkNormalizeToken of module vtkNormalize"
    Err.Number = VTK_UNEXPECTED_ERROR
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkNormalizeString
' Author    : Lucas Vitorino
' Purpose   : Normalize a String by normalizing the VBA identifier tokens in it
'               - an identifier token is a String starting by a [A-Za-z] character with nothing but characters,
'                 numbers and underscores in it
'               - comments are not scanned for identifier tokens.
' Returns   : The normalized String corresponding to the input String.
' Raises    : - VTK_UNEXPECTED_EOS
'             - VTK_UNEXPECTED_CHAR
'             - VTK_UNEXPECTED_ERROR
' Notes     : This code is based on code generated by Klemen's LEX4VB. Get LEX4VB from http://www.schmidks.de
'---------------------------------------------------------------------------------------
'
Public Function vtkNormalizeString(s As String, listOfTokens() As String) As String

Dim token As String
Dim State As Integer, OldState As Integer
Dim Cnt As Integer
Dim ch As String
Dim p As Integer

Dim returnString As String

On Error GoTo vtkNormalizeString_Error
p = 1: State = 0: OldState = -1
s = s & Chr(0)
    
Do While p <= Len(s)
    If State = OldState Then Cnt = Cnt + 1 Else Cnt = 0
    OldState = State
    ch = Mid(s, p, 1)
        Select Case State
            Case 0:
                ' The analyser is looking for a token : copy characters without modifying
                If Asc(ch) = 0 Then
                    State = 9
                ElseIf ch Like "[A-Za-z]" Then
                    token = token & ch
                    State = 1
                ElseIf ch Like "[']" Then
                    returnString = returnString & ch
                    token = ""
                    State = 2
                ElseIf ch Like "[""]" Then
                    returnString = returnString & ch
                    token = ""
                    State = 3
                ElseIf ch Like "[!']" Then
                    returnString = returnString & ch
                    token = ""
                    State = 0
                Else: Err.Raise VTK_UNEXPECTED_CHAR
                End If

            Case 1:
                ' The analsyer is in a token : normalize tokens it finds
                If Asc(ch) = 0 Then
                    returnString = returnString & vtkNormalizeToken(token, listOfTokens)
                    State = 9
                ElseIf ch Like "[A-Za-z,0-9,_]" Then
                    token = token & ch
                    State = 1
                ElseIf ch Like "[']" Then
                    returnString = returnString & vtkNormalizeToken(token, listOfTokens) & ch
                    token = ""
                    State = 2
                ElseIf ch Like "[""]" Then
                    returnString = returnString & vtkNormalizeToken(token, listOfTokens) & ch
                    token = ""
                    State = 3
                ElseIf ch Like "[!']" Then
                    returnString = returnString & vtkNormalizeToken(token, listOfTokens) & ch
                    token = ""
                    State = 0
                Else: Err.Raise VTK_UNEXPECTED_CHAR
                End If


            Case 2:
                ' The analyser is in a comment : copy characters without modifying until end of String
                If Asc(ch) = 0 Then
                    State = 9
                ElseIf Asc(ch) > 0 Then
                    returnString = returnString & ch
                    State = 2
                Else: Err.Raise VTK_UNEXPECTED_CHAR
                End If
                
            Case 3:
                'The analyser is in a String : copy characters without modifying until a quote
                If Asc(ch) = 0 Then
                    State = 9
                ElseIf ch Like "[""]" Then
                    returnString = returnString & ch
                    State = 0
                ElseIf Asc(ch) > 0 Then
                    returnString = returnString & ch
                    State = 3
                Else: Err.Raise VTK_UNEXPECTED_CHAR
                End If
                    

            Case 9:
                If True Then
                    State = 9
                Else: Err.Raise VTK_UNEXPECTED_CHAR
                End If

        End Select

p = p + 1
Loop

If State <> 9 Then Err.Raise VTK_UNEXPECTED_EOS

vtkNormalizeString = returnString
Exit Function

vtkNormalizeString_Error:
    
    Err.Source = "function vtkNormalizeString of module vtkNormalize"
    
    Select Case Err.Number
        Case VTK_UNEXPECTED_EOS
            Err.Description = "Unexpected EOS in String " & s
        Case VTK_UNEXPECTED_CHAR
            Err.Description = "Unexpected character of Ascii code " & Asc(ch) & " in String " & s & " at position " & p
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkNormalizeFile
' Author    : Lucas Vitorino
' Purpose   : Normalize a file.
'               - Create a temporary file in the same directory
'               - Copy each line of the original file in the temporary file, after calling vtkNormalizeString on them
'               - Delete the original file
'               - Rename the temporary file with the name of the original one.
' Raises    : - VTK_WRONG_FILE_PATH
'             - VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Sub vtkNormalizeFile(fullFilePath As String, listOfTokens() As String)

On Error GoTo vtkNormalizeFile_Error

    Dim fso As New FileSystemObject
    
    ' Initialize input and output files
    Dim inputFileObject As File
    Set inputFileObject = fso.GetFile(fullFilePath)
    
    Dim normalizedFullFilePath As String
    normalizedFullFilePath = inputFileObject.ParentFolder & "\" & "tmp_" & inputFileObject.name
    fso.CreateTextFile (normalizedFullFilePath)
    Dim outputFileObject As File
    Set outputFileObject = fso.GetFile(normalizedFullFilePath)
    
    ' Initialize objects to read and write the files
    Dim textFileRead As TextStream
    Set textFileRead = fso.OpenTextFile(fullFilePath, ForReading)
    Dim textFileWrite As TextStream
    Set textFileWrite = fso.OpenTextFile(normalizedFullFilePath, ForWriting)
    
    ' Copy each line of the input file in the output file after normalizing it
    Do Until textFileRead.AtEndOfStream
        textFileWrite.WriteLine (vtkNormalizeString(textFileRead.ReadLine, listOfTokens))
    Loop
    
    ' Close the streams
    textFileRead.Close
    textFileWrite.Close
    
    ' Delete original file
    Kill fullFilePath
    
    ' Rename normalized file with the name of the original file
    outputFileObject.name = fso.GetFileName(fullFilePath)
    
   On Error GoTo 0
   
   Exit Sub

vtkNormalizeFile_Error:

    Err.Source = "sub vtkNormalizeFile of module vtkNormalize"
    
    Select Case Err.Number
        Case 53
            Err.Number = VTK_WRONG_FILE_PATH
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
    Exit Sub
    
End Sub


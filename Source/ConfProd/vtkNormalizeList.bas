Attribute VB_Name = "vtkNormalizeList"
'---------------------------------------------------------------------------------------
' Module    : vtkNormalizeList
' Author    : Lucas Vitorino
' Purpose   : This module contains the list of words to be normalized.
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

Private listOfWordsToNormalize As String

Private Sub initializeList()
    listOfWordsToNormalize = _
    "Dim" & "," & _
    "Wb" & "," & _
    "Err" & "," & _
    "File" & "," & _
    "Folder" & "," & _
    "Scripting" & "," & _
    ""
End Sub


'---------------------------------------------------------------------------------------
' Procedure : vtkListOfWordsToNormalize
' Author    : Lucas Vitorino
' Purpose   : This functions initializes the array containing the properly cased Strings.
'---------------------------------------------------------------------------------------
'
Public Function vtkListOfWordsToNormalize() As String()
    initializeList
    vtkListOfWordsToNormalize = Split(listOfWordsToNormalize, ",")
End Function

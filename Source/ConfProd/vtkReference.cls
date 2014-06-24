VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkReference
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : Contains informations relative to a VTK Reference
'               - 'ID' is the ID of the reference for XML exportation
'               - 'name' of the reference
'               - 'GUID' of the reference
'               - 'relPath' path of the reference relative to the AppData user's folder
'                   or fullPath if it's outside of the AppData user's folder
'               - 'fulPath' of the reference
' Method    :
'               - initWithVBAReference
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
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

 Private m_id As String    ' Mandatory ID of the reference
 Private m_name As String  ' Mandatory for each reference
 Private m_guid As String  ' guid and path are exclusive each other
 Private m_relPath As String

'---------------------------------------------------------------------------------------
' Property  : ID as String
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : The ID of the reference for XML exportation
'---------------------------------------------------------------------------------------
'
Property Get id() As String
    id = m_id
End Property

Property Let id(new_id As String)
    m_id = new_id
End Property

'---------------------------------------------------------------------------------------
' Property  : name as String
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : The name of the reference
'---------------------------------------------------------------------------------------
'
Property Get name() As String
    name = m_name
End Property

Property Let name(new_name As String)
    m_name = new_name
End Property

'---------------------------------------------------------------------------------------
' Property  : GUID as String
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : The GUID of the reference
'---------------------------------------------------------------------------------------
'
Property Get GUID() As String
    GUID = m_guid
End Property

Property Let GUID(new_GUID As String)
    m_guid = new_GUID
End Property

'---------------------------------------------------------------------------------------
' Property  : relPath as String
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : The path of the reference relative to the AppData user's folder
'---------------------------------------------------------------------------------------
'
Property Get relPath() As String
    relPath = m_relPath
End Property

Property Let relPath(new_relPath As String)
    m_relPath = new_relPath
End Property

'---------------------------------------------------------------------------------------
' Property  : fullPath as String
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : The path of the reference relative to the AppData user's folder
'---------------------------------------------------------------------------------------
'
Property Get fullPath() As String
    fullPath = convertRelativePathToFullPath(m_relPath)
End Property

Property Let fullPath(new_fullPath As String)
    m_relPath = convertFullPathToRelativePath(new_fullPath)
End Property

'---------------------------------------------------------------------------------------
' Function  : convertRelativePathToFullPath
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : convert a relative reference path to a full path into the AppData user's folder
'---------------------------------------------------------------------------------------
'
Private Function convertRelativePathToFullPath(relPath As String) As String
    If vtkIsPathAbsolute(relPath) Then
        convertRelativePathToFullPath = relPath
       Else
        convertRelativePathToFullPath = Environ("AppData") & "\" & relPath
    End If
End Function

'---------------------------------------------------------------------------------------
' Function  : convertFullPathToRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : convert a full reference path to a path relative to the AppData user's folder
'             - doesn't change the full path if it's outside the AppData user's folder
'---------------------------------------------------------------------------------------
'
Private Function convertFullPathToRelativePath(fullPath As String) As String
    convertFullPathToRelativePath = Replace(fullPath, Environ("AppData") & "\", "")
End Function

'---------------------------------------------------------------------------------------
' Subroutine : initWithVBAReference
' Author     : Jean-Pierre Imbert
' Date       : 14/06/2014
' Purpose    : initialize the instance with a VBA reference
'---------------------------------------------------------------------------------------
'
Public Sub initWithVBAReference(ref As Reference)
    Me.name = ref.name
    Me.GUID = ref.GUID
    Me.fullPath = ref.fullPath
End Sub

'---------------------------------------------------------------------------------------
' Subroutine : addToWorkbook
' Author     : Jean-Pierre Imbert
' Date       : 22/06/2014
' Purpose    : add the vtkReference to the workbook given as a parameter
'              - no error if the reference is already set
'---------------------------------------------------------------------------------------
'
Public Sub addToWorkbook(Wb As Workbook)
    On Error Resume Next
    If Me.GUID = "" Then
        Wb.VBProject.references.AddFromFile Me.fullPath
       Else
        Wb.VBProject.references.AddFromGuid Me.GUID, 0, 0
    End If
    On Error GoTo 0
End Sub

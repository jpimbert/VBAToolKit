VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkReference
' Author    : Lucas Vitorino
' Purpose   : Manage the attributes of a reference to ease our life when recreating configurations.
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

Private m_name As String
Private m_guid As String
Private m_path As String

'---------------------------------------------------------------------------------------
' Procedure : name
' Author    : Lucas Vitorino
' Purpose   : Get the name of the current vtkReference.
'---------------------------------------------------------------------------------------
'
Public Property Get name() As String
    name = m_name
End Property

'---------------------------------------------------------------------------------------
' Procedure : name
' Author    : Lucas Vitorino
' Purpose   : Set the name of the current vtkReference
'---------------------------------------------------------------------------------------
'
Public Property Let name(sname As String)
    m_name = sname
End Property

'---------------------------------------------------------------------------------------
' Procedure : guid
' Author    : Lucas Vitorino
' Purpose   : Get the guid of the current vtkReference.
'---------------------------------------------------------------------------------------
'
Public Property Get guid() As String
    guid = m_guid
End Property

'---------------------------------------------------------------------------------------
' Procedure : guid
' Author    : Lucas Vitorino
' Purpose   : Set the guid of the current vtkReference.
'---------------------------------------------------------------------------------------
'
Public Property Let guid(sguid As String)
    m_guid = sguid
End Property

'---------------------------------------------------------------------------------------
' Procedure : path
' Author    : Lucas Vitorino
' Purpose   : Get the path of the current vtkReference.
'---------------------------------------------------------------------------------------
'
Public Property Get path() As String
    path = m_path
End Property

'---------------------------------------------------------------------------------------
' Procedure : path
' Author    : Lucas Vitorino
' Purpose   : Set the path of the current vtkReference.
'---------------------------------------------------------------------------------------
'
Public Property Let path(spath As String)
    m_path = spath
End Property

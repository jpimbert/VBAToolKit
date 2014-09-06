VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Module    : vtkProject
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Manage a VB Toolkit Project and its attributes
'
'   Fonctions à implémenter
'
' - Trouver le fichier Excel ouvert qui contient ce nom de projet (et le mettre en cache)
' - Attacher les fonctions de chemin à cet objet
' - Attacher cet objet et le configuration manager associé
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

Private m_projectName As String

'---------------------------------------------------------------------------------------
' Property  : projectName As String
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : This property must be set for the object
'---------------------------------------------------------------------------------------
'
Public Property Get projectName() As String
    projectName = m_projectName
End Property

Public Property Let projectName(ByVal sprojectName As String)
    m_projectName = sprojectName
End Property

'---------------------------------------------------------------------------------------
' Method    : initialized As Boolean (Private)
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return true if the vtkProject object is initialized
'               - the projectName property must nt be null
'---------------------------------------------------------------------------------------
'
Private Function initialized() As Boolean
    initialized = (m_projectName <> "")
End Function

'---------------------------------------------------------------------------------------
' Property  : projectDEVName As String (Read Only)
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return the name of the development project, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get projectDEVName() As String
    If initialized Then
        projectDEVName = m_projectName & "_DEV"
       Else
        projectDEVName = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : projectDEVStandardRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return the relative path of the development project, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get projectDEVStandardRelativePath() As String
    If initialized Then
        projectDEVStandardRelativePath = "Project\" & workbookDEVName
       Else
        projectDEVStandardRelativePath = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : XMLConfigurationStandardRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return the relative path of the XML configuration file, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get XMLConfigurationStandardRelativePath() As String
    If initialized Then
        XMLConfigurationStandardRelativePath = "Project\" & XMLConfigurationName
       Else
        XMLConfigurationStandardRelativePath = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : projectStandardRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Return the relative path of the delivery project, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get projectStandardRelativePath() As String
    If initialized Then
        projectStandardRelativePath = "Delivery\" & workbookName
       Else
        projectStandardRelativePath = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : workbookDEVName
' Author    : Jean-Pierre Imbert
' Date      : 05/06/2013
' Purpose   : Return the full name of the development workbook, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get workbookDEVName() As String
    If initialized Then
        workbookDEVName = projectDEVName & vtkDefaultExcelExtension
       Else
        workbookDEVName = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : workbookName
' Author    : Jean-Pierre Imbert
' Date      : 05/06/2013
' Purpose   : Return the full name of the delivery workbook, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get workbookName() As String
    If initialized Then
        workbookName = projectName & vtkDefaultExcelExtension
       Else
        workbookName = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : XMLConfigurationName
' Author    : Jean-Pierre Imbert
' Date      : 13/10/2013
' Purpose   : Return the name of the XML configuration file, or null if the object is not initialized
'---------------------------------------------------------------------------------------
'
Public Property Get XMLConfigurationName() As String
    If initialized Then
        XMLConfigurationName = projectName & ".xml"
       Else
        XMLConfigurationName = ""
    End If
End Property


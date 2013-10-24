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
Private m_projectBeforeRootFolderPath As String
Private m_xmlRelativeFolderPath As String

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
        workbookDEVName = projectDEVName & ".xlsm"
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
        workbookName = projectName & ".xlsm"
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

'---------------------------------------------------------------------------------------
' Procedure : projectBeforeRootFolderPath
' Author    : Lucas Vitorino
' Purpose   : Set the absolute path of the root folder containing the project folder.
'---------------------------------------------------------------------------------------
'
Public Property Let projectBeforeRootFolderPath(ByVal folderPath As String)
    m_projectBeforeRootFolderPath = folderPath
End Property

'---------------------------------------------------------------------------------------
' Procedure : xmlRelativeFolderPath
' Author    : Lucas Vitorino
' Purpose   : Set the the relative path of the folder in which the xml sheet is created.
'---------------------------------------------------------------------------------------
'
Public Property Let xmlRelativeFolderPath(ByVal folderPath As String)
    m_xmlRelativeFolderPath = folderPath
End Property

'---------------------------------------------------------------------------------------
' Procedure : projectRootFolderPath
' Author    : Lucas Vitorino
' Purpose   : Return the absolute path of the root folder of the project.
'             - If the name is not set, return "".
'             - If the name is set but no other path has been set, return the $testFolderPath \ $projectName
'             - If the name is set and the before root folder path has been set, $beforeRootFolderPath \ $projectname
'---------------------------------------------------------------------------------------
'
Public Property Get projectRootFolderPath() As String
    If initialized Then
        Dim fso As New FileSystemObject
        If m_projectBeforeRootFolderPath <> "" Then
            projectRootFolderPath = fso.BuildPath(m_projectBeforeRootFolderPath, projectName)
        Else
            projectRootFolderPath = fso.BuildPath(VBAToolKit.vtkTestPath, projectName)
        End If
    Else
        projectRootFolderPath = ""
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : projectBeforeRootFolderPath
' Author    : Lucas Vitorino
' Purpose   : Return the absolute path of the "before root" folder of the project.
'             - If the name is not set, return "".
'             - If the name is set but no other path has been set, return the $testFolderPath
'             - If the name is set and the before root folder path has been set, $beforeRootFolderPath
'---------------------------------------------------------------------------------------
'
Public Property Get projectBeforeRootFolderPath() As String
    If initialized Then
        If m_projectBeforeRootFolderPath <> "" Then
            projectBeforeRootFolderPath = m_projectBeforeRootFolderPath
        Else
            projectBeforeRootFolderPath = VBAToolKit.vtkTestPath
        End If
    Else
        projectBeforeRootFolderPath = ""
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : xmlRelativeFolderPath
' Author    : Lucas Vitorino
' Purpose   : Return the relative path of the folder where the xml sheet is stored.
'             If the path has not been set, return "".
'---------------------------------------------------------------------------------------
'
Public Property Get xmlRelativeFolderPath() As String
    If m_xmlRelativeFolderPath <> "" Then
        xmlRelativeFolderPath = m_xmlRelativeFolderPath
    Else
        xmlRelativeFolderPath = ""
    End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : xmlSheetFullPath
' Author    : Lucas Vitorino
' Purpose   : Return the full path of the xml sheet, made of
'               - the absolute path of the project's root folder
'               - the relative path of the project xml folder
'               - the name of the xml sheet ($projectName.xml)
'             If the name or the root path are
'---------------------------------------------------------------------------------------
'
Public Property Get xmlSheetFullPath() As String
    If initialized And (projectRootFolderPath <> "") Then
        Dim fso As New FileSystemObject
        xmlSheetFullPath = fso.BuildPath(fso.BuildPath(projectRootFolderPath, xmlRelativeFolderPath), XMLConfigurationName)
    Else
        xmlSheetFullPath = ""
    End If
End Property

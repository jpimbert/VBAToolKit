VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_projectName As String

'---------------------------------------------------------------------------------------
' Module    : vtkProject
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   : Manage a VB Toolkit Project and its attributes
'---------------------------------------------------------------------------------------

'
'   Fonctions à implémenter
'
' - Trouver le fichier Excel ouvert qui contient ce nom de projet (et le mettre en cache)
' - Attacher les fonctions de chemin à cet objet
' - Attacher cet objet et le configuration manager associé
'

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



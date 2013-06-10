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
' Property  : projectDEVName AS String (ead only)
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get projectDEVName() As String
    If m_projectName = "" Then
        projectDEVName = ""
       Else
        projectDEVName = m_projectName & "_DEV"
    End If
End Property

'---------------------------------------------------------------------------------------
' Property  : projectDEVStandardRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get projectDEVStandardRelativePath() As String
    projectDEVStandardRelativePath = "Project\" & workbookDEVName
End Property

'---------------------------------------------------------------------------------------
' Property  : projectStandardRelativePath
' Author    : Jean-Pierre Imbert
' Date      : 04/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get projectStandardRelativePath() As String
    projectStandardRelativePath = "Delivery\" & workbookname
End Property

'---------------------------------------------------------------------------------------
' Property  : workbookDEVName
' Author    : Jean-Pierre Imbert
' Date      : 05/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get workbookDEVName() As String
    workbookDEVName = projectDEVName & ".xlsm"
End Property


'---------------------------------------------------------------------------------------
' Property  : workbookName
' Author    : Jean-Pierre Imbert
' Date      : 05/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get workbookname() As String
    workbookname = projectName & ".xlsm"
End Property

'---------------------------------------------------------------------------------------
' Procedure : DEVProjectPath
' Author    : user
' Date      : 09/06/2013
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ProjectFullPath() As String
Dim fso As New FileSystemObject
    ProjectFullPath = fso.GetParentFolderName(Workbooks(workbookDEVName()).path)
End Property

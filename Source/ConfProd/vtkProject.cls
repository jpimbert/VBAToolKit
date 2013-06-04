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

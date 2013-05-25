VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfigurationManager
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Manage Configurations and Modules of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a Module can be imported/exported to a specific path relative to the VTK Project directory
'               - each configuration is a VBA project included in a specific Excel file whose path is relative to the VTK Project directory
'
' Usage:
'   - Each instance of Configuration Manager is attached to an Excel Worksheet (supposed to be a VTK project)
'       - the method vtkConfigurationManagerForWorkbook of module vtkConfigurationManagers give the instance attached to a workbook, or create it
'   - a VTK project is implemented by a main Excel file containing a sheet named "vtkConfigurations"
'       - if such a sheet doesn't exist, it is created
'
'---------------------------------------------------------------------------------------

Private m_workbookName As String
Private m_configurationSheet As Worksheet
Private Const sheetName As String = "vtkConfigurations"

Private Sub Class_Terminate()
    Set m_configurationSheet = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Property  : workbookName as String - Read/Write - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : name of the project attached to this configuration manager
'               - when changed, the configurationWorkSheet private property is seset
'---------------------------------------------------------------------------------------
'
Public Property Get workbookName() As String
    workbookName = m_workbookName
End Property

Public Property Let workbookName(ByVal sworkbookName As String)
    If m_workbookName <> sworkbookName Then
        m_workbookName = sworkbookName
        setConfigurationSheet      ' Reset the configuration sheet when the attached project is changed
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : setConfigurationSheet
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : set then configurationSheet property
'               - Create and initialize the sheet if it doesn't exist in the Project file
'---------------------------------------------------------------------------------------
'
Private Sub setConfigurationSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Workbooks(workbookName).Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then        ' If the Configuration doesn't exist, create one
        Set ws = Workbooks(workbookName).Worksheets.Add(Type:=xlWorksheet)
        ws.name = sheetName
        End If
    Set m_configurationSheet = ws
End Sub

'---------------------------------------------------------------------------------------
' Procedure : configurationSheet as Worksheet - Read only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Return the configuration Sheet - used during tests
'---------------------------------------------------------------------------------------
'
Public Property Get configurationSheet() As Worksheet
    Set configurationSheet = m_configurationSheet
End Property

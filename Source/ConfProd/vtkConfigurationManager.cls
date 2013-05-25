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
' To implement later
'   - automatic recognition of the configuration sheet version, and proposal to the user to convert to the last version
'       cell $A$1 contains the version identification of the sheet version like "vtkConfigurations v1.0"
'
'---------------------------------------------------------------------------------------

Private m_workbookName As String
Private m_configurationSheet As Worksheet
Private Const sheetName As String = "vtkConfigurations"
Private Const currentVersion As String = "vtkConfigurations v1.0"

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
        '   Initialize the configuration sheet
        initializeConfigurationSheet ws
        End If
    ' Verify that the configuration sheet is a
    If ws.Range("A1") <> currentVersion Then
        ' Code to convert from previous version will be implemented here
        End If
    Set m_configurationSheet = ws
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initializeConfigurationSheet
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Initialize a new configuration sheet
'---------------------------------------------------------------------------------------
'
Private Sub initializeConfigurationSheet(ws As Worksheet)
    ws.Range("A1") = currentVersion
    ws.Range("A2") = "Module Name"
    ws.Range("B1") = Workbooks(workbookName).VBProject.name
    ws.Range("B2") = "Delivery\" & workbookName     ' Default init for test purpose, will be initialized later
    ws.Range("C1") = Workbooks(workbookName).VBProject.name & "_DEV"
    ws.Range("C2") = "Project\" & workbookName & "_DEV"   ' Default init for test purpose, will be initialized later
    
    ws.Columns("A:C").ColumnWidth = 22
    ws.Range("A1").Font.Size = 8
    ws.Range("A2").HorizontalAlignment = xlCenter
    ws.Range("A2").Font.Bold = True
    ws.Range("B1:C2").HorizontalAlignment = xlCenter
    ws.Range("B1:C1").Font.Bold = True
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

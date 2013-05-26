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
    '   Set the initialized worksheet property
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

'---------------------------------------------------------------------------------------
' Property  : configurationCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Get the number of configurations
'---------------------------------------------------------------------------------------
'
Public Property Get configurationCount() As Integer
    ' The number of configurations is the number of used columns minus one
    configurationCount = configurationSheet.UsedRange.Columns.Count - 1
End Property

'---------------------------------------------------------------------------------------
' Function  : configuration
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : return the n-th configuration of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function configuration(n As Integer) As String
    Dim sResult As String
    sResult = ""
    If n >= 1 And n <= configurationCount Then
        sResult = configurationSheet.Cells(1, n + 1)
       Else
        sResult = ""
    End If
    configuration = sResult
End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationNumber
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the number of a configuration given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationNumber(configuration As String) As Integer
    Dim r As Range
    On Error Resume Next
    Set r = configurationSheet.Rows(1)      ' Select first column of the sheet
    Set r = r.Offset(0, 1)                  ' Forget the two first rows
    Set r = r.Find(what:=configuration, SearchOrder:=xlByRows)
    On Error GoTo 0
    If r Is Nothing Then
        getConfigurationNumber = 0
       Else
        getConfigurationNumber = r.Column - 1
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : getConfigurationPathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : return the n-th relative configuration path of the project, given n as integer
'               - return "" if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPathWithNumber(n As Integer) As String
    Dim sResult As String
    sResult = ""
    If n >= 1 And n <= configurationCount Then
        sResult = configurationSheet.Cells(2, n + 1)
       Else
        sResult = ""
    End If
    getConfigurationPathWithNumber = sResult
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationPathWithNumber
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : change the n-th relative configuration path of the project, given n as integer
'               - does nothing if the configuration is inexistant
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationPathWithNumber(n As Integer, path As String)
    If n >= 1 And n <= configurationCount Then
        configurationSheet.Cells(2, n + 1) = path
    End If
End Sub

'---------------------------------------------------------------------------------------
' Function  : getConfigurationPath
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the relative path of a configuration given its name
'               - return "" of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getConfigurationPath(configuration As String) As String
    getConfigurationPath = getConfigurationPathWithNumber(getConfigurationNumber(configuration))
End Function

'---------------------------------------------------------------------------------------
' Procedure : setConfigurationPath
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : change the path of a configuration given its name
'---------------------------------------------------------------------------------------
'
Public Sub setConfigurationPath(configuration As String, path As String)
    setConfigurationPathWithNumber n:=getConfigurationNumber(configuration), path:=path
End Sub

'---------------------------------------------------------------------------------------
' Property  : moduleCount as Integer - Read Only - Public
' Author    : Jean-Pierre Imbert
' Date      : 25/05/2013
' Purpose   : Get the number of modules
'---------------------------------------------------------------------------------------
'
Public Property Get moduleCount() As Integer
    ' The number of modules is the number of used rows minus two
    moduleCount = configurationSheet.UsedRange.Rows.Count - 2
End Property

'---------------------------------------------------------------------------------------
' Function  : module
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the n-th module of the project, given n as integer
'               - return "" if the module is inexistant
'---------------------------------------------------------------------------------------
'
Public Function module(n As Integer) As String
    Dim sResult As String
    sResult = ""
    If n >= 1 And n <= moduleCount Then
        sResult = configurationSheet.Cells(n + 2, 1)
       Else
        sResult = ""
    End If
    module = sResult
End Function

'---------------------------------------------------------------------------------------
' Procedure : getModuleNumber
' Author    : Jean-Pierre Imbert
' Date      : 26/05/2013
' Purpose   : return the number of a module given its name
'               - return 0 of the name is inexistant
'---------------------------------------------------------------------------------------
'
Public Function getModuleNumber(module As String) As Integer
    Dim r As Range
    On Error Resume Next
    Set r = configurationSheet.Columns(1)   ' Select first column of the sheet
    Set r = r.Offset(2, 0)                  ' Forget the two first rows
    Set r = r.Find(what:=module, SearchOrder:=xlByColumns)
    On Error GoTo 0
    If r Is Nothing Then
        getModuleNumber = 0
       Else
        getModuleNumber = r.Row - 2
    End If
End Function




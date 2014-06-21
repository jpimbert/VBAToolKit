VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkReferenceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkReferenceManager
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   : Manage Configuration References of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a set of VBA references can be set
'
' Usage:
'   - Each instance of Reference Manager is attached to the configuration manager of a VTK project
'       - this attachment is managed by the configuration manager, the reference manager doesn't know it
'   - a VTK project is implemented by a main Excel file containing a sheet named "vtkReferences"
'       - if such a sheet doesn't exist, it is created
'       - the Excel file in which this sheet must be created is given by the configuration manager
'         at the reference manager initialisation
'       - configuration names in the vtkReferences sheet are linked to the vtkConfigurations sheet
'   - When initializing an instance, the relative development project workbook must be opened
'       - else initialization is not performed
'   - The robustness of the class is delegated to the conf manager
'       - It's supposed, and mandatory, that a vtkConfigurations sheet exists in the workbook
'       - the class method are never called when the class is not initialized with a consistent workbook
'       - no check is performed on the confNumber value
'   - The refManager could be re-initialized and its behavior is consistent
'   - The relations between configurations and references are managed by hand in the vtkReferences sheet,
'     the refManager is only a reader
'
' Methods:
'   - init(Wb,confCount) initialize or re-initialize the refManager instance
'   - referenceSheet, readonly property for test usage
'   - addConfiguration add a column at the end of the sheet
'   - allReferences() gives the collection of all references attached to a configuration or not
'       - A reference is communicated with the vtkReference type
'   - references(confNumber) gives the collection of references attached to the configuration
'       - A reference is communicated with the vtkReference type
'
' Note:
'   - References path are absolute or relative to the App Data user's folder (Environ("AppData"))
'
' To implement later
'   - automatic recognition of the references sheet version, and proposal to the user to convert to the last version
'       cell $A$1 contains the version identification of the sheet version like "vtkReferences v1.0"
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

Private m_workbook As Workbook
Private m_referenceSheet As Worksheet
Private m_nbTitleColumnsInConfSheet As Integer

Private Const sheetName As String = "vtkReferences"
Private Const currentVersion As String = "vtkReferences v1.0"
Private Const nbTitleColumns As Integer = 3
Private Const nbTitleRows As Integer = 1

Private Sub Class_Terminate()
    Set m_referenceSheet = Nothing
    Set m_workbook = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Subroutine  : init
' Author      : Jean-Pierre Imbert
' Date        : 10/06/2014
' Purpose     : Initialize the reference manager. called by the configuration manager
' Parameters  : - Workbook of the DEV configuration for the project, will contain the reference sheet
'               - Configuration count in the configuration sheet
'               - Title columns count in the configuration sheet
'               - optional devConf number
'                   the devConf is attached to VBAToolKit reference when the reference sheet is created
'                   if devConf = 0, there is no devConf, and VBAToolKit is not attached, but included
'                   if devConf is <0 or >confCount, no attachment is performed
'---------------------------------------------------------------------------------------
'
Public Sub init(Wb As Workbook, confCount As Integer, nbTitleColumnsInConfSheet As Integer, Optional devConf As Integer = -1)
    If Not m_workbook Is Wb Then
        Set m_workbook = Wb
        m_nbTitleColumnsInConfSheet = nbTitleColumnsInConfSheet
        If Not m_workbook Is Nothing Then
            setReferenceSheet confCount, devConf    ' Reset the reference sheet when the attached workbook is changed
           Else
            Set m_referenceSheet = Nothing
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Property  : referenceSheet, workSheet readonly
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   : Get the vtkReferences workSheet (for test purpose)
'---------------------------------------------------------------------------------------
'
Public Property Get referenceSheet() As Worksheet
    Set referenceSheet = m_referenceSheet
End Property

'---------------------------------------------------------------------------------------
' Subroutine  : setReferenceSheet, private
' Author      : Jean-Pierre Imbert
' Date        : 10/06/2014
' Purpose     : - Set the referenceSheet readonly property
'               - Create the vtkReferences sheet if it doesn't exist in the workbook
'                 - initialize the reference list with a default one
'               - Verify the configuration count in the vtkReferences sheet
'                 - and fix it if inconsistent
'               - Verify that all references in the workbook are listed in the vtkReferences sheet
'                 - and fix it if inconsistent
' Parameter   : - confCount as Integer, count of configurations
'               - devConf as integer, number of the development configuration
'                   the devConf is attached to VBAToolKit reference when the reference sheet is created
'                   if devConf = 0, there is no devConf, and VBAToolKit is not attached, but included
'                   if devConf is <0 or >confCount, no attachment is performed
'---------------------------------------------------------------------------------------
'
Private Sub setReferenceSheet(confCount As Integer, devConf As Integer)
    On Error Resume Next
    Set m_referenceSheet = m_workbook.Worksheets(sheetName)
    On Error GoTo 0
    If m_referenceSheet Is Nothing Then        ' If the Configuration doesn't exist, create one
        '   Set the initialized worksheet property
        Set m_referenceSheet = m_workbook.Worksheets.Add(Type:=xlWorksheet)
        m_referenceSheet.name = sheetName
        '   Initialize the reference sheet
        initializeReferenceSheet confCount, devConf
        End If
    ' Verify that the configuration sheet is in last known version
    If m_referenceSheet.Range("A1") <> currentVersion Then
        ' Code to convert from previous version will be implemented here or in a specific method
        End If
    '   Update the configurations in the reference sheet
    alignConfigurationCount confCount
    '   Update the references list in the reference sheet
    addReferences
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initializeReferenceSheet
' Author    : Jean-Pierre Imbert
' Date      : 11/06/2014
' Purpose   : Initialize a new reference sheet
'             Attach the standard references to all configuration (except VBAToolKit, attached to the Dev Conf)
' Parameter : - confCount as Integer, count of configurations
'             - devConf as integer, number of the development configuration
'                   the devConf is attached to VBAToolKit reference when the reference sheet is created
'                   if devConf = 0, there is no devConf, and VBAToolKit is not attached, but included
'                   if devConf is <0 or >confCount, no attachment is performed
'---------------------------------------------------------------------------------------
'
Private Sub initializeReferenceSheet(confCount As Integer, devConf As Integer)
    ' Init titles
    m_referenceSheet.Range("A1") = currentVersion
    m_referenceSheet.Range("A1").Font.Size = 8
    m_referenceSheet.Range("B1") = ""
    m_referenceSheet.Range("C1") = "-"      ' addConfiguration works with the usedRange property
    m_referenceSheet.Columns("A:C").ColumnWidth = 22
    ' Init standard references
    m_referenceSheet.Range("A2") = "VBA"
    m_referenceSheet.Range("B2") = "{000204EF-0000-0000-C000-000000000046}"
    m_referenceSheet.Range("A3") = "Excel"
    m_referenceSheet.Range("B3") = "{00020813-0000-0000-C000-000000000046}"
    m_referenceSheet.Range("A4") = "stdole"
    m_referenceSheet.Range("B4") = "{00020430-0000-0000-C000-000000000046}"
    m_referenceSheet.Range("A5") = "Office"
    m_referenceSheet.Range("B5") = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
    m_referenceSheet.Range("A6") = "MSForms"
    m_referenceSheet.Range("B6") = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
    m_referenceSheet.Range("A7") = "Scripting"
    m_referenceSheet.Range("B7") = "{420B2830-E718-11CF-893D-00A0C9054228}"
    m_referenceSheet.Range("A8") = "VBIDE"
    m_referenceSheet.Range("B8") = "{0002E157-0000-0000-C000-000000000046}"
    m_referenceSheet.Range("A9") = "Shell32"
    m_referenceSheet.Range("B9") = "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}"
    m_referenceSheet.Range("A10") = "MSXML2"
    m_referenceSheet.Range("B10") = "{F5078F18-C551-11D3-89B9-0000F81FE221}"
    m_referenceSheet.Range("A11") = "ADODB"
    m_referenceSheet.Range("B11") = "{00000206-0000-0010-8000-00AA006D2EA4}"
    m_referenceSheet.Range("A12") = "VBAToolKit"
    m_referenceSheet.Range("C12") = "Microsoft\AddIns\VBAToolKit.xlam"
    ' Init configuration titles
    Dim i As Integer
    For i = 1 To confCount
        addConfiguration
        Dim j As Integer
        For j = 1 To referenceCount
            If devConf >= 0 And devConf <= confCount Then m_referenceSheet.Cells(nbTitleRows + j, nbTitleColumns + i) = "X"
            If m_referenceSheet.Cells(nbTitleRows + j, 1) = "VBAToolKit" And i <> devConf Then m_referenceSheet.Cells(nbTitleRows + j, nbTitleColumns + i) = ""
        Next j
    Next i
End Sub

'---------------------------------------------------------------------------------------
' Procedure : alignConfigurationCount
' Author    : Jean-Pierre Imbert
' Date      : 11/06/2014
' Purpose   : align the reference sheet with the configuration Count
'             - add Columns if necesary
'---------------------------------------------------------------------------------------
'
Private Sub alignConfigurationCount(confCount As Integer)
    If confCount > configurationCount Then  ' we have to add columns
        Dim i As Integer
        For i = configurationCount + 1 To confCount
            addConfiguration
        Next i
    End If
End Sub

'---------------------------------------------------------------------------------------
' Function  : configurationCount
' Author    : Jean-Pierre Imbert
' Date      : 11/06/2014
' Purpose   : Configuration count of the references sheet
'---------------------------------------------------------------------------------------
'
Private Function configurationCount() As Integer
    configurationCount = m_referenceSheet.UsedRange.Columns.Count - nbTitleColumns
End Function

'---------------------------------------------------------------------------------------
' Function  : referenceCount
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : Reference count of the references sheet
'---------------------------------------------------------------------------------------
'
Private Function referenceCount() As Integer
    referenceCount = m_referenceSheet.UsedRange.Rows.Count - nbTitleRows
End Function

'---------------------------------------------------------------------------------------
' Subroutine  : addConfiguration
' Author      : Jean-Pierre Imbert
' Date        : 10/06/2014
' Purpose     : Add a Configuration
'               - doesn't add references to the new configuration
'               - exit silently if called with a refManager not initialized
'---------------------------------------------------------------------------------------
'
Public Sub addConfiguration()
    Dim newColumn As Integer
    If Not m_referenceSheet Is Nothing Then
        newColumn = m_referenceSheet.UsedRange.Columns.Count + 1
        m_referenceSheet.Columns(newColumn).ColumnWidth = 22
        m_referenceSheet.Columns(newColumn).HorizontalAlignment = xlCenter
        m_referenceSheet.Cells(1, newColumn).FormulaLocal = "=INDIRECT(ADRESSE(1;" & newColumn - nbTitleColumns + m_nbTitleColumnsInConfSheet & ";4;1;""vtkConfigurations""))"
        m_referenceSheet.Cells(1, newColumn).Font.Bold = True
       Else
        Err.Raise Number:=VTK_NOTINITIALIZED, Source:="vtkReferenceManager:addConfiguration", Description:="Impossible to add a configuration to the reference sheet of a null workbook."
    End If
End Sub

'---------------------------------------------------------------------------------------
' Subroutine  : addReferences
' Author      : Jean-Pierre Imbert
' Date        : 14/06/2014
' Purpose     : Add references contained in the attached workbook to the reference sheet
'---------------------------------------------------------------------------------------
'
Private Sub addReferences()
    Dim ref As Reference, ra As Range, newRow As Integer, vtkRef As New vtkReference
    For Each ref In m_workbook.VBProject.references
        Set ra = m_referenceSheet.Columns("A:A").Find(ref.name, , xlValues, xlWhole)
        If ra Is Nothing Then
            vtkRef.initWithVBAReference ref
            newRow = m_referenceSheet.UsedRange.Rows.Count + 1
            m_referenceSheet.Cells(newRow, 1) = vtkRef.name
            If ref.GUID Like "" Then
                m_referenceSheet.Cells(newRow, 3) = vtkRef.relPath
               Else
                m_referenceSheet.Cells(newRow, 2) = vtkRef.GUID
            End If
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Function  : allReferences as collection
' Author    : Jean-Pierre Imbert
' Date      : 14/06/2014
' Purpose   : Get all the references attached to a configuration or not
' Return    : a collection of vtkReference type
'---------------------------------------------------------------------------------------
'
Public Function allReferences() As Collection
    Set allReferences = references(0)
End Function

'---------------------------------------------------------------------------------------
' Function  : references as collection
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   : Get the references attached to a configuration
' Return    : a collection of vtkReference type
' Parameter : - confNumber, the number of the configuration from which to get the references collection
'               if confNumber = 0, allReferences are returned
'               if confNumber > conf count, VTK_FORBIDDEN_PARAMETER Error is returned
'---------------------------------------------------------------------------------------
'
Public Function references(confNumber As Integer) As Collection
    If Not m_referenceSheet Is Nothing Then
        If Not (confNumber < 0 Or confNumber > configurationCount) Then
            Dim c As New Collection, i As Integer, r As vtkReference, s As String, b As Boolean
            ' Loop for each reference
            For i = 1 To referenceCount
                ' Create a new vtkReference instance
                Set r = New vtkReference
                r.id = "r" & i
                r.name = m_referenceSheet.Cells(nbTitleRows + i, 1)
                r.GUID = m_referenceSheet.Cells(nbTitleRows + i, 2)
                r.relPath = m_referenceSheet.Cells(nbTitleRows + i, 3)
                ' Evaluate b as True if the erference must be added to the collection
                If confNumber = 0 Then
                    b = True
                   Else
                    s = m_referenceSheet.Cells(nbTitleRows + i, nbTitleColumns + confNumber)
                    Select Case s
                       Case "X"
                        b = True
                       Case ""
                        b = False
                       Case Else
                        Err.Raise Number:=VTK_INVALID_FIELD, Source:="vtkReferenceManager:references", Description:="Field (" & s & ") must be blank or X in the reference sheet (ref=" & r.name & ", conf=" & m_referenceSheet.Cells(1, nbTitleColumns + confNumber) & ") for the workbook " & m_workbook.name
                    End Select
                End If
                ' Add the reference to the collection
                If b Then c.Add Item:=r, Key:=r.id
            Next i
            Set references = c
           Else
            Err.Raise Number:=VTK_FORBIDDEN_PARAMETER, Source:="vtkReferenceManager:references", Description:="confNumber (" & confNumber & ") is out of bound (0," & configurationCount & ") for the workbook " & m_workbook.name
        End If
       Else
        Err.Raise Number:=VTK_NOTINITIALIZED, Source:="vtkReferenceManager:references", Description:="Impossible to get references from a sheet of a null workbook"
    End If
End Function
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
'   - references(confNumber) gives the collection of references attached to the configuration
'       - A reference is communicated with the vtkReference type
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
' Purpose     : Set the workbook
'---------------------------------------------------------------------------------------
'
Public Sub init(Wb As Workbook, confCount As Integer)
    If Not m_workbook Is Wb Then
        Set m_workbook = Wb
        If Not m_workbook Is Nothing Then setReferenceSheet (confCount)     ' Reset the reference sheet when the attached workbook is changed
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
    referenceSheet = m_referenceSheet
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
' Parameter   : - confCount as Integer, number of configurations
'---------------------------------------------------------------------------------------
'
Private Sub setReferenceSheet(confCount As Integer)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = m_workbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then        ' If the Configuration doesn't exist, create one
        Set ws = m_workbook.Worksheets.Add(Type:=xlWorksheet)
        ws.name = sheetName
        '   Initialize the configuration sheet
'        initializeReferenceSheet ws
        End If
    ' Verify that the configuration sheet is in last known version
    If ws.Range("A1") <> currentVersion Then
        ' Code to convert from previous version will be implemented here or in a specific method
        End If
    '   Set the initialized worksheet property
    Set m_referenceSheet = ws
End Sub

'---------------------------------------------------------------------------------------
' Subroutine  : addConfiguration
' Author      : Jean-Pierre Imbert
' Date        : 10/06/2014
' Purpose     : Add a Configuration
'---------------------------------------------------------------------------------------
'
Public Sub addConfiguration()

End Sub

'---------------------------------------------------------------------------------------
' Function  : references as collection
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   : Get the references attached to a configuration
' Return    : a collection of vtkReference type
' Parameter : - confNumber, the number of the configuration from which to get the references collection
'---------------------------------------------------------------------------------------
'
Public Function references(confNumber As Integer) As Collection

End Function

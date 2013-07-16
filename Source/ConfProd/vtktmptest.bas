Attribute VB_Name = "vtktmptest"
'Option Explicit

Private vtkButtons As New Collection

Public Function testfn()
a = vtkCreateProject(vtkTestPath, "testtest")

C = vtkExportAll("VBAToolKit.xlsm")
f = vtkImportTestConfig()
dd = vtkImportTestConfig()
Debug.Print f
End Function

Public Sub AddToolBar()
    Dim v As CommandBar, b As CommandBarControl, e As vtkButtonEvent
    ' Delete previous toolbar
    On Error Resume Next
    Application.VBE.CommandBars("VBAToolKit").Delete
    On Error GoTo 0
    ' Create and configure Toolbar
    Set v = Application.VBE.CommandBars.Add(name:="VBAToolKit", Position:=msoBarFloating)
    v.Visible = True
    ' Create and configure Button
    Set b = v.Controls.Add(Type:=msoControlButton)
    b.Caption = "Create Project"
    b.OnAction = "CreateProjectClicked"
    b.FaceId = 2031
    ' create and configure Button Event
    Set e = New vtkButtonEvent
    Set e.commandButton = b
    vtkButtons.Add e
End Sub

Public Sub CreateProjectClicked()
    Debug.Print "Bouton cliqué"
End Sub


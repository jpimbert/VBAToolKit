Attribute VB_Name = "vtkToolsBars"
'---------------------------------------------------------------------------------------
' Module    : vtkToolBar
' Author    : user
' Date      : 06/06/2013
' Purpose   : - create commandbar in IDE how allow to create project
'             - Add New Module
'---------------------------------------------------------------------------------------


Private VtkEvtHandlers As New VtkEventHandlers

'---------------------------------------------------------------------------------------
' Procedure : CreateToolsBarAndButton
' Author    : user
' Date      : 06/06/2013
' Purpose   : -create command bar in ide
'---------------------------------------------------------------------------------------
'
Sub CreateToolsBarAndButton()
    'Name of the cammandbar can not be like the project name
    Const cCommandBar = "VbaToolKit_Bar"

    Dim bar As CommandBar
    Dim creprj As CommandBarControl
    Dim evh As VtkEventHandler

    ' Delete all Event Handlers
    VtkEvtHandlers.Clear

    ' Delete the Commandbar if it already exists
    For Each bar In Application.VBE.CommandBars
        If bar.name = cCommandBar Then bar.Delete
    Next
    Set bar = Application.VBE.CommandBars.Add(name:=cCommandBar, Position:=msoBarTop, Temporary:=True)

   ' Create Project Function Button
    Set creprj = bar.Controls.Add(Type:=msoControlButton)
    With creprj
        .FaceId = 2031
        .Caption = "Create Project"
        .TooltipText = "Click here to create a new project"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Create_Project", creprj
    End With
        
    bar.Visible = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Create_Project
' Author    : user
' Date      : 07/06/2013
' Purpose   : - call form how allo to write project name and path
'---------------------------------------------------------------------------------------
'
Sub Create_Project()
    vtkCreateProjectForm.Show
End Sub



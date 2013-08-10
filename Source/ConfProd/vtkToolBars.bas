Attribute VB_Name = "vtkToolBars"
'---------------------------------------------------------------------------------------
' Module    : vtkToolBars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Module for Toolbars and ToolbarControls management
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : toolBarName
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Give the Excel and VBE ToolBar name
'             - The Toolbars name is derived from the Current Project name
'             - both toolbars share the same name
'---------------------------------------------------------------------------------------
'
Private Function toolBarName() As String
    toolBarName = ThisWorkbook.VBProject.name
End Function

'---------------------------------------------------------------------------------------
' Procedure : controlTag
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Give the tag used for the specific VBAToolKit controls
'             - the tag are used to get these controls with the FindControls method
'             - The tag is derived from the Current Project Name
'---------------------------------------------------------------------------------------
'
Private Function controlTag() As String
    controlTag = ThisWorkbook.VBProject.name & "_Tag"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateToolbars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Create Excel and VBE toolbars
'             - Don't recreate them if they already exists
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateToolbars()
    Dim barE As CommandBar, BarV As CommandBar
    
    ' Create Excel Commandbar if necessary
    On Error Resume Next
    Set barE = Application.CommandBars(toolBarName)
    On Error GoTo 0
    If barE Is Nothing Then Set barE = Application.CommandBars.Add(name:=toolBarName, Position:=msoBarFloating)
    barE.Visible = True
    
        ' Create VBE Commandbar if necessary
    On Error Resume Next
    Set BarV = Application.VBE.CommandBars(toolBarName)
    On Error GoTo 0
    If BarV Is Nothing Then Set BarV = Application.VBE.CommandBars.Add(name:=toolBarName, Position:=msoBarTop)
    BarV.Visible = True
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkDeleteToolbars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Delete Excel and VBE Toolbars
'---------------------------------------------------------------------------------------
'
Public Sub vtkDeleteToolbars()
   On Error Resume Next
    ' Delete Excel Commandbar if necessary
    Application.CommandBars(toolBarName).Delete
    ' Delete VBE Commandbar if necessary
    Application.VBE.CommandBars(toolBarName).Delete
   On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
'' Procedure : CreateToolsBarAndButton
'' Author    : user
'' Date      : 06/06/2013
'' Purpose   : -create command bar in ide
''---------------------------------------------------------------------------------------
''
'Sub CreateToolsBarAndButton()
'    'Name of the cammandbar can not be like the project name
'    Const cCommandBar = "VbaToolKit_Bar"
'
'    Dim Bar As CommandBar
'    Dim CrePrj As CommandBarControl
'    Dim GitStat As CommandBarControl
'    Dim AddMod As CommandBarControl
'
'    Dim evh As vtkEventHandler
'
'    ' Delete all Event Handlers
'   VtkEvtHandlers.Clear
'
'    ' Delete the Commandbar if it already exists
'    For Each Bar In Application.VBE.CommandBars
'        If Bar.name = cCommandBar Then
'        Bar.Delete
'        End If
'    Next
'
'   ' Create a new Command Bar in vbe
'    Set Bar = Application.VBE.CommandBars.Add(name:=cCommandBar, Position:=msoBarTop, Temporary:=True)
'
'   ' Create Project Function Button
'    Set CrePrj = Bar.Controls.Add(Type:=msoControlButton)
'    ' gitstatus Function Button
'    Set GitStat = Bar.Controls.Add(Type:=msoControlButton)
'    ' addmodule Function Button
'    Set AddMod = Bar.Controls.Add(Type:=msoControlButton)
'
'    With CrePrj
'        .FaceId = 2031
'        .Caption = "Create Project"
'        .TooltipText = "Click here to create a new project"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "Create_Project", CrePrj
'    End With
'
'        With GitStat
'        .FaceId = 49
'        .Caption = "Git Status"
'        .TooltipText = "Click here to show git status"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "Git_Status", GitStat
'    End With
'
'        With AddMod
'        .FaceId = 2520
'        .Caption = "add module"
'        .TooltipText = "Click here to add new module"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "AddModule", AddMod
'    End With
'
'    Bar.Visible = True
'
'End Sub
'Sub CreateExcelBarAndButton()
'    'Name of the cammandbar can not be like the project name
'   Const XlCommandBar = "VbaToolKit_Bar"
'
'    Dim BarEx As CommandBar
'    Dim UpdateVbeToolbarEx As CommandBarButton
'    Dim CrePrjEx As CommandBarButton
'    Dim GitStatEx As CommandBarButton
'    Dim AddmodEx As CommandBarButton
'
'    Dim evh As vtkEventHandler
'
'    ' Delete all Event Handlers
'    VtkEvtHandlers.Clear
'
'    ' Delete the Commandbar if it already exists
'    For Each BarEx In CommandBars
'        If BarEx.name = XlCommandBar Then
'        BarEx.Delete
'        End If
'    Next
'
'    ' Create a new Command Bar in excel
'    Set BarEx = CommandBars.Add(name:=XlCommandBar, Position:=msoBarFloating)
'
'
'   ' Add button 1 to this bar
'    Set CrePrjEx = BarEx.Controls.Add(Type:=msoControlButton)
'   ' Add button 2 to this bar
'    Set GitStatEx = BarEx.Controls.Add(Type:=msoControlButton)
'   ' Add button 3 to this bar
'    Set AddmodEx = BarEx.Controls.Add(Type:=msoControlButton)
'   ' Add button 1 to this bar
'    Set UpdateVbeToolbarEx = BarEx.Controls.Add(Type:=msoControlButton)
'
'
'    With CrePrjEx
'        .FaceId = 2031
'        .Caption = "Create Project"
'        .TooltipText = "Click here to create a new project"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "Create_Project", CrePrjEx
'    End With
'
'    With GitStatEx
'        .FaceId = 49
'        .Caption = "Git Status"
'        .TooltipText = "Click here to show git status"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "Git_Status", GitStatEx
'    End With
'        With AddmodEx
'        .FaceId = 2520
'        .Caption = "add module"
'        .TooltipText = "Click here to add new module"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "AddModule", AddmodEx
'    End With
'
'        With UpdateVbeToolbarEx
'        .FaceId = 37
'        .Caption = "Update VBE Buttons"
'        .TooltipText = "Click here to Update VBE Buttons"
'        .Style = msoButtonIconAndCaption
'        VtkEvtHandlers.AddNew "Update_VBE_Buttons", UpdateVbeToolbarEx
'    End With
'
'    BarEx.Visible = True
'End Sub
'
''---------------------------------------------------------------------------------------
'' Procedure : Create_Project
'' Author    : user
'' Date      : 07/06/2013
'' Purpose   : - call form how allo to write project name and path
''---------------------------------------------------------------------------------------
''
'Sub Create_Project()
'    vtkCreateProjectForm.Show
'End Sub
'Sub Git_Status()
'   Dim RetValGitStatus As String
'   RetValGitStatus = vtkStatusGit()
'   MsgBox (RetValGitStatus)
'End Sub
'Sub AddModule()
'  VtkAddModule.Show
'End Sub
'
'Sub Update_VBE_Buttons()
'    Call CreateToolsBarAndButton
'End Sub


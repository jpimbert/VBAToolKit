Attribute VB_Name = "vtkToolBars"
'---------------------------------------------------------------------------------------
' Module    : vtkToolBars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Module for Toolbars and ToolbarControls management
'---------------------------------------------------------------------------------------
Option Explicit

Private colEventHandlers As Collection
Private buttonClicked As Boolean ' for test with dummy callback

'---------------------------------------------------------------------------------------
' Procedure : vtkAddEventHandler
' Author    : Jean-Pierre Imbert
' Date      : 20/08/2013
' Purpose   : Create an event handler and associates it with a button and action.
'---------------------------------------------------------------------------------------
'
Public Sub vtkAddEventHandler(action As String, cmdBarCtl As CommandBarControl)
    Dim evh As New vtkEventHandler
    cmdBarCtl.onAction = action
    Set evh.cbe = Application.VBE.Events.CommandBarEvents(cmdBarCtl)
    If colEventHandlers Is Nothing Then Set colEventHandlers = New Collection
    colEventHandlers.Add Item:=evh, Key:=action
    Set evh = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkClearEventHandlers
' Author    : Jean-Pierre Imbert
' Date      : 20/08/2013
' Purpose   : Delete all event handlers associated to the VBE toolbar
'---------------------------------------------------------------------------------------
'
Public Sub vtkClearEventHandlers()
    Set colEventHandlers = Nothing
End Sub

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
    Dim barE As CommandBar, BarV As CommandBar, cbControl As CommandBarControl
    
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

' Faire une fonction pour ajouter un bouton (à tester avec une fonction Mock de test (+ variable privée)
' à tester dans les deux barres d'outils

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
    Application.CommandBars(toolBarName).Delete     ' Unit tests confirm that the button are deleted with the toolbar
    ' Delete VBE Commandbar if necessary
    Application.VBE.CommandBars(toolBarName).Delete
    ' Delete Event Handlers for VBE
    vtkClearEventHandlers
   On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateToolbarButton
' Author    : Jean-Pierre Imbert
' Date      : 20/08/2013
' Purpose   : Create a button in both toolbars, giving
'               - tha caption of the button
'               - the help text of the button
'               - the faceId of the button (see http://fring.developpez.com/vba/excel/faceid/)
'               - the name of the procedure to call when the button is clicked
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateToolbarButton(caption As String, helpText As String, faceId As Integer, action As String)
    Dim cbControl As CommandBarButton

        ' Create Button in the Excel Command Bar
    Set cbControl = Application.CommandBars(toolBarName).Controls.Add(Type:=msoControlButton)
    cbControl.faceId = faceId
    cbControl.caption = caption
    cbControl.TooltipText = helpText
    cbControl.Style = msoButtonAutomatic
    cbControl.onAction = action

        ' Create Same Button in the VBE Command Bar
    Set cbControl = Application.VBE.CommandBars(toolBarName).Controls.Add(Type:=msoControlButton)
    cbControl.faceId = faceId
    cbControl.caption = caption
    cbControl.TooltipText = helpText
    cbControl.Style = msoButtonAutomatic
    vtkAddEventHandler action:=action, cmdBarCtl:=cbControl

'    Set cbControl = Application.VBE.CommandBars(toolBarName).Controls.Add(Type:=msoControlButton)
'    cbControl.faceId = 2031
'    cbControl.caption = "Create Project"
'    cbControl.TooltipText = "Click here to create a new project"
'    cbControl.Style = msoButtonAutomatic
'    vtkAddEventHandler action:="vtkTestButtonClick", cmdBarCtl:=cbControl
    
End Sub

'
'---------------------------------------------------------------------------------------
' Dummy callback for button event test
'---------------------------------------------------------------------------------------
'
Public Sub vtkTestCommandBarButtonClicked()
    buttonClicked = True
End Sub
Public Sub vtkTestCommandBarButtonClickedReset()
    buttonClicked = False
End Sub
Public Function vtkIsTestCommandBarButtonClicked() As Boolean
    vtkIsTestCommandBarButtonClicked = buttonClicked
End Function

'Public Sub vtkTestButtonClick()
'    Debug.Print "Button Clicked"
'End Sub

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


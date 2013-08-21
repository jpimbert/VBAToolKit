Attribute VB_Name = "vtkToolBars"
'---------------------------------------------------------------------------------------
' Module    : vtkToolBars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Module for Toolbars and ToolbarControls management
'
' TODO : add button for VBE ToolKit reactivation (FaceId = 688)
'        add button for Add a Module (FaceId = 2646)
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
' Purpose   : Create Complete Excel and VBE toolbars
'             - Don't recreate them if they already exists
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateToolbars()
    vtkCreateEmptyToolbars
    vtkCreateToolbarButton caption:="Create Project", helpText:="Click here to create a new project", faceId:=2031, action:="VBAToolKit.vtkShowCreateProjectForm"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateEmptyToolbars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Create Empty (no buttons) Excel and VBE toolbars
'             - Don't recreate them if they already exists
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateEmptyToolbars()
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

'
'---------------------------------------------------------------------------------------
' Procedures : Call back for buttons
' Author    : Jean-Pierre Imbert
' Date      : 20/08/2013
' Purpose   : Just a wrapper to forms, or unit tested subs
'---------------------------------------------------------------------------------------
'
Private Sub vtkShowCreateProjectForm()
    vtkCreateProjectForm.Show
End Sub

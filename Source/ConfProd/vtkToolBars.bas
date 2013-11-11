Attribute VB_Name = "vtkToolBars"
'---------------------------------------------------------------------------------------
' Module    : vtkToolBars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Module for Toolbars and ToolbarControls management
'
' TODO :
'        add button for Add a Module (FaceId = 2646)
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
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
' Procedure : projectName
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Give name of running project
'---------------------------------------------------------------------------------------
'
Private Function projectName() As String
    projectName = ThisWorkbook.VBProject.name
End Function

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
    toolBarName = projectName
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
    controlTag = projectName & "_Tag"
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateToolbars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Create Complete Excel and VBE toolbars
'             - Don't recreate them if they already exists
' Params    :
'             - vbeToolbar, boolean true (default) if the VBE Toolbar must be created
'             - excToolbar, boolean true (default) if the Excel Toolbar must be created
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateToolbars(Optional vbeToolbar As Boolean = True, Optional excToolbar As Boolean = True)
    vtkCreateEmptyToolbars vbeToolbar:=vbeToolbar, excToolbar:=excToolbar
    ' Create the button for VBE Toolbar reactivation
    If excToolbar Then
        vtkCreateToolbarButton caption:="Reset VBE Toolbar", _
                               helpText:="Click here to reset the VBA IDE Toolbar", _
                               faceId:=688, _
                               action:=projectName & ".vtkReactivateVBEToolBar", _
                               vbeToolbar:=False, _
                               excToolbar:=True
    End If
                               
    ' Create other buttons
    vtkCreateToolbarButton caption:="Create Project", _
                           helpText:="Click here to create a new project", _
                           faceId:=2031, _
                           action:=projectName & ".vtkShowCreateProjectForm", _
                           vbeToolbar:=vbeToolbar, _
                           excToolbar:=excToolbar

    ' Create the button for recreate configuration
    vtkCreateToolbarButton caption:="Recreate Configuration", _
                           helpText:="Click here to recreate a configuration", _
                           faceId:=680, _
                           action:=projectName & ".vtkShowRecreateConfigurationForm", _
                           vbeToolbar:=vbeToolbar, _
                           excToolbar:=excToolbar
                           
                           
End Sub

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateEmptyToolbars
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Create Empty (no buttons) Excel and VBE toolbars
'             - Don't recreate them if they already exists
' Params    :
'             - vbeToolbar, boolean true (default) if the VBE Toolbar must be created
'             - excToolbar, boolean true (default) if the Excel Toolbar must be created
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateEmptyToolbars(Optional vbeToolbar As Boolean = True, Optional excToolbar As Boolean = True)
    Dim barE As CommandBar, barV As CommandBar, cbControl As CommandBarControl
    
    ' Create Excel Commandbar if necessary
    If excToolbar Then
        On Error Resume Next
        Set barE = Application.CommandBars(toolBarName)
        On Error GoTo 0
        If barE Is Nothing Then Set barE = Application.CommandBars.Add(name:=toolBarName, Position:=msoBarFloating)
        barE.Visible = True
    End If
    
    ' Create VBE Commandbar if necessary
    If vbeToolbar Then
        On Error Resume Next
        Set barV = Application.VBE.CommandBars(toolBarName)
        On Error GoTo 0
        If barV Is Nothing Then Set barV = Application.VBE.CommandBars.Add(name:=toolBarName, Position:=msoBarTop)
        barV.Visible = True
    End If
    
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
' Procedure : vtkReactivateVBEToolBar
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   : Recreate all event handlers for VBE Toolbar after VBA reset
'---------------------------------------------------------------------------------------
'
Public Sub vtkReactivateVBEToolBar()
    Dim c As CommandBarControl
    vtkClearEventHandlers
    For Each c In Application.VBE.CommandBars(toolBarName).Controls
        vtkAddEventHandler action:=c.onAction, cmdBarCtl:=c
    Next
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
'               - vbeToolbar, boolean true (default) if the VBE Toolbar must be modified
'               - excToolbar, boolean true (default) if the Excel Toolbar must be modified
'---------------------------------------------------------------------------------------
'
Public Sub vtkCreateToolbarButton(caption As String, helpText As String, faceId As Integer, action As String, Optional vbeToolbar As Boolean = True, Optional excToolbar As Boolean = True)
    Dim cbControl As CommandBarButton

        ' Create Button in the Excel Command Bar
    If excToolbar Then
        Set cbControl = Application.CommandBars(toolBarName).Controls.Add(Type:=msoControlButton)
        cbControl.faceId = faceId
        cbControl.caption = caption
        cbControl.TooltipText = helpText
        cbControl.Style = msoButtonAutomatic
        cbControl.onAction = action
    End If

        ' Create Same Button in the VBE Command Bar
    If vbeToolbar Then
        Set cbControl = Application.VBE.CommandBars(toolBarName).Controls.Add(Type:=msoControlButton)
        cbControl.faceId = faceId
        cbControl.caption = caption
        cbControl.TooltipText = helpText
        cbControl.Style = msoButtonAutomatic
        vtkAddEventHandler action:=action, cmdBarCtl:=cbControl
    End If

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

Public Sub vtkShowRecreateConfigurationForm()
    vtkRecreateConfigurationForm.Show
End Sub

' Special CallBack, manually configured in Excel to recreate AddIn
'   Options Excel, Personnaliser, choisir Macro pour bouton de raccourci rapide
Public Sub vtkClickForVBAToolKitRecreation()
    vtkRecreateConfiguration projectName:="VBAToolKit", configurationName:="VBAToolKit"
End Sub

' Special CallBack, manually configured in Excel to recreate Dev Project
'   Must be run from another project named VBAToolKit2
Public Sub vtkClickForVBAToolKitDEVRecreation()
    If Not ActiveWorkbook.VBProject.name Like "VBAToolKit2_DEV" Then Exit Sub
    If Not ActiveWorkbook.name Like "VBAToolKit2_DEV.xlsm" Then Exit Sub
    vtkRecreateConfiguration projectName:="VBAToolKit2", configurationName:="VBAToolKit_DEV"
End Sub



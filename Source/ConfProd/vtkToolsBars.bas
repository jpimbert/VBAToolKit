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
   
    Dim Bar As CommandBar
    Dim CrePrj As CommandBarControl
    Dim GitStat As CommandBarControl
    Dim AddMod As CommandBarControl
   
    Dim evh As VtkEventHandler

    ' Delete all Event Handlers
   VtkEvtHandlers.Clear

    ' Delete the Commandbar if it already exists
    For Each Bar In Application.VBE.CommandBars
        If Bar.name = cCommandBar Then
        Bar.Delete
        End If
    Next
   
   ' Create a new Command Bar in vbe
    Set Bar = Application.VBE.CommandBars.Add(name:=cCommandBar, Position:=msoBarTop, Temporary:=True)

   ' Create Project Function Button
    Set CrePrj = Bar.Controls.Add(Type:=msoControlButton)
    ' gitstatus Function Button
    Set GitStat = Bar.Controls.Add(Type:=msoControlButton)
    ' addmodule Function Button
    Set AddMod = Bar.Controls.Add(Type:=msoControlButton)
    
    With CrePrj
        .FaceId = 2031
        .Caption = "Create Project"
        .TooltipText = "Click here to create a new project"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Create_Project", CrePrj
    End With
        
        With GitStat
        .FaceId = 49
        .Caption = "Git Status"
        .TooltipText = "Click here to show git status"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Git_Status", GitStat
    End With
    
        With AddMod
        .FaceId = 2520
        .Caption = "add module"
        .TooltipText = "Click here to add new module"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "AddModule", AddMod
    End With
            
    Bar.Visible = True

End Sub
Sub CreateExcelBarAndButton()
    'Name of the cammandbar can not be like the project name
   Const XlCommandBar = "VbaToolKit_Bar"
    
    Dim BarEx As CommandBar
    Dim UpdateVbeToolbarEx As CommandBarButton
    Dim CrePrjEx As CommandBarButton
    Dim GitStatEx As CommandBarButton
    Dim AddmodEx As CommandBarButton
    
    Dim evh As VtkEventHandler

    ' Delete all Event Handlers
    VtkEvtHandlers.Clear
   
    ' Delete the Commandbar if it already exists
    For Each BarEx In CommandBars
        If BarEx.name = XlCommandBar Then
        BarEx.Delete
        End If
    Next
    
    ' Create a new Command Bar in excel
    Set BarEx = CommandBars.Add(name:=XlCommandBar, Position:=msoBarFloating)


   ' Add button 1 to this bar
    Set CrePrjEx = BarEx.Controls.Add(Type:=msoControlButton)
   ' Add button 2 to this bar
    Set GitStatEx = BarEx.Controls.Add(Type:=msoControlButton)
   ' Add button 3 to this bar
    Set AddmodEx = BarEx.Controls.Add(Type:=msoControlButton)
   ' Add button 1 to this bar
    Set UpdateVbeToolbarEx = BarEx.Controls.Add(Type:=msoControlButton)
  
 
    With CrePrjEx
        .FaceId = 2031
        .Caption = "Create Project"
        .TooltipText = "Click here to create a new project"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Create_Project", CrePrjEx
    End With
    
    With GitStatEx
        .FaceId = 49
        .Caption = "Git Status"
        .TooltipText = "Click here to show git status"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Git_Status", GitStatEx
    End With
        With AddmodEx
        .FaceId = 2520
        .Caption = "add module"
        .TooltipText = "Click here to add new module"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "AddModule", AddmodEx
    End With
    
        With UpdateVbeToolbarEx
        .FaceId = 37
        .Caption = "Update VBE Buttons"
        .TooltipText = "Click here to Update VBE Buttons"
        .Style = msoButtonIconAndCaption
        VtkEvtHandlers.AddNew "Update_VBE_Buttons", UpdateVbeToolbarEx
    End With

    BarEx.Visible = True
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
Sub Git_Status()
   Dim RetValGitStatus As String
   RetValGitStatus = vtkStatusGit()
   MsgBox (RetValGitStatus)
End Sub
Sub AddModule()
  VtkAddModule.Show
End Sub

Sub Update_VBE_Buttons()
    Call CreateToolsBarAndButton
End Sub

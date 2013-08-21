VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ThisWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Manage the toolbars during install/uninstall of addin
'             - This module exists in slightly different Prod and Test versions
'               - explicit call to VBAToolKit project for Prod version
'               - explicit call to VBAToolKit_DEV project for Test version
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub Workbook_AddinInstall()
    VBAToolKit.vtkToolBars.vtkCreateToolbars
End Sub

Private Sub Workbook_AddinUninstall()
    VBAToolKit.vtkToolBars.vtkDeleteToolbars
End Sub

Private Sub Workbook_Open()
    If AddIns("VBAToolKit").Installed Then VBAToolKit.vtkToolBars.vtkCreateToolbars excToolbar:=False
End Sub

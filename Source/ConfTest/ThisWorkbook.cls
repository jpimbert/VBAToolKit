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
'
' WARNING :
'      This module exists in slightly different Prod and Test versions
'          - explicit call to VBAToolKit project for Prod version
'          - explicit call to VBAToolKit_DEV project for Test version
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

Private Sub Workbook_AddinInstall()
    VBAToolKit_DEV.vtkToolBars.vtkCreateToolbars
End Sub

Private Sub Workbook_AddinUninstall()
    VBAToolKit_DEV.vtkToolBars.vtkDeleteToolbars
End Sub

' JUST FOR TEST PURPOSES
'-----------------------
Private Sub Workbook_Open()
    VBAToolKit.vtkToolBars.vtkDeleteToolbars
    VBAToolKit_DEV.vtkToolBars.vtkCreateToolbars
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    VBAToolKit_DEV.vtkToolBars.vtkDeleteToolbars
End Sub
'-----------------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    VBAToolKit.vtkExportConfiguration ThisWorkbook.VBProject, "VBAToolKit", "VBAToolKit_DEV", True
End Sub



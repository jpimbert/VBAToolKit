VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkApplicationEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _
    ByVal SaveAsUI As Boolean, Cancel As Boolean)
        
        Debug.Print "Saving " & Wb.name
        Cancel = False

End Sub



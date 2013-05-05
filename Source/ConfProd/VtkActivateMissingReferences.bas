Attribute VB_Name = "VtkActivateMissingReferences"
'---------------------------------------------------------------------------------------
' Module    : VtkAvtivateReferences
' Author    : user
' Date      : 26/04/2013
' Purpose   : - check that workrbook is open and activate vbide and +-scripting references
'---------------------------------------------------------------------------------------


Option Explicit

Public Function VtkActivateReferences(workbookname As String)

If VtkWorkbokIsOpen(workbookname) = False Then     'if the workbook is open

On Error Resume Next ' if the first extention is already activated, we will try to activate the second one

Workbooks(workbookname).VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 0, 0  ' +- to activate Scripting : Microsoft scripting runtime
Workbooks(workbookname).VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0 ' to activate VBIDE: Microsoft visual basic for applications extensibility 5.3

End If

End Function


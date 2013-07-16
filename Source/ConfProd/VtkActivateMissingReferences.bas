Attribute VB_Name = "VtkActivateMissingReferences"
'---------------------------------------------------------------------------------------
' Module    : VtkActivateReferences
' Author    : Abdelfattah Lahbib
' Date      : 26/04/2013
' Purpose   : - Check that workbook is open, activate VB IDE and +-scripting references
'---------------------------------------------------------------------------------------


Option Explicit

Public Function VtkActivateReferences(workbookName As String)

If VtkWorkbookIsOpenFunction(workbookName) = True Then     'if the workbook is open

On Error Resume Next ' if the first extention is already activated, we try to activate the second one

Workbooks(workbookName).VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 0, 0  ' +- to activate Scripting : Microsoft scripting runtime
Workbooks(workbookName).VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0 ' to activate VB IDE: Microsoft visual basic for applications extensibility 5.3

End If

End Function


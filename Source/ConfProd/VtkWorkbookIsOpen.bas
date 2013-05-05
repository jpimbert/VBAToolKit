Attribute VB_Name = "VtkWorkbookIsOpen"
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : VtkWorkbokIsOpen
' Author    : user
' Date      : 24/04/2013
' Purpose   :- this function return true if the workbook is already open , and false if it's close
'---------------------------------------------------------------------------------------
'
Public Function VtkWorkbokIsOpen(workbookname As String) As Boolean 'to test it debug.Print a=VtkWorkbokIsOpen("VBAToolKit")

 On Error Resume Next
 Workbooks(workbookname).Activate 'if we have a problem to activate workbook = the workbook is closed , if we can activate it without problem = the workbook is open
    If Err <> 0 Then
        VtkWorkbokIsOpen = True
    Else
        VtkWorkbokIsOpen = False
    End If
End Function


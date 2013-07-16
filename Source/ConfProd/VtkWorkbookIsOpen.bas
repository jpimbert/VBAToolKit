Attribute VB_Name = "VtkWorkbookIsOpen"
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : VtkWorkbookIsOpen
' Author    : Abdelfattah Lahbib
' Date      : 24/04/2013
' Purpose   : - Return true if the Workbook is already open , false if it's closed
'---------------------------------------------------------------------------------------
'
Public Function VtkWorkbookIsOpenFunction(workbookName As String) As Boolean
'to test it : debug.Print a=VtkWorkbokIsOpen("VBAToolKit")

 On Error GoTo Err_function
 Workbooks(workbookName).Activate
 'if we have a problem to activate the workbook = it is closed,
 'if we can activate it without problem = the workbook is open
    
Err_function:
    If Err = 0 Then
        VtkWorkbookIsOpenFunction = True
    Else
        VtkWorkbookIsOpenFunction = False
    End If
End Function


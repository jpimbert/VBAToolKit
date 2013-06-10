Attribute VB_Name = "VtkWorkbookIsOpen"

Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : VtkWorkbookIsOpen
' Author    : user
' Date      : 24/04/2013
' Purpose   :- this function return true if the workbook is already open , and false if it's close
'---------------------------------------------------------------------------------------
'
Public Function VtkWorkbookIsOpenFunction(workbookname As String) As Boolean 'to test it debug.Print a=VtkWorkbokIsOpen("VBAToolKit")

 On Error GoTo Err_function
 Workbooks(workbookname).Activate 'if we have a problem to activate workbook = the workbook is closed , if we can activate it without problem = the workbook is open
    
Err_function:
    If Err = 0 Then
        VtkWorkbookIsOpenFunction = True
    Else
        VtkWorkbookIsOpenFunction = False
    End If
End Function



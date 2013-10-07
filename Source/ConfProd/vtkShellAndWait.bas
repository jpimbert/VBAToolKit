Attribute VB_Name = "vtkShellAndWait"
Option Explicit
Option Compare Text

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modShellAndWait
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' This page on the web site: www.cpearson.com/Excel/ShellAndWait.aspx
' 9-September-2008
'
' This module contains code for the ShellAndWait function that will Shell to a process
' and wait for that process to end before returning to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000

Public Enum ShellAndWaitResult
    Success = 0
    Failure = 1
    TimeOut = 2
    InvalidParameter = 3
    SysWaitAbandoned = 4
    UserWaitAbandoned = 5
    UserBreak = 6
End Enum

Public Enum ActionOnBreak
    IgnoreBreak = 0
    AbandonWait = 1
    PromptUser = 2
End Enum

Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const STATUS_WAIT_0 As Long = &H0
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_INFINITE = -1&


Public Function ShellAndWait(ShellCommand As String, _
                    TimeOutMs As Long, _
                    ShellWindowState As VbAppWinStyle, _
                    BreakKey As ActionOnBreak) As ShellAndWaitResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShellAndWait
'
' This function calls Shell and passes to it the command text in ShellCommand. The function
' then waits for TimeOutMs (in milliseconds) to expire.
'
'   Parameters:
'       ShellCommand
'           is the command text to pass to the Shell function.
'
'       TimeOutMs
'           is the number of milliseconds to wait for the shell'd program to wait. If the
'           shell'd program terminates before TimeOutMs has expired, the function returns
'           ShellAndWaitResult.Success = 0. If TimeOutMs expires before the shell'd program
'           terminates, the return value is ShellAndWaitResult.TimeOut = 2.
'
'       ShellWindowState
'           is an item in VbAppWinStyle specifying the window state for the shell'd program.
'
'       BreakKey
'           is an item in ActionOnBreak indicating how to handle the application's cancel key
'           (Ctrl Break). If BreakKey is ActionOnBreak.AbandonWait and the user cancels, the
'           wait is abandoned and the result is ShellAndWaitResult.UserWaitAbandoned = 5.
'           If BreakKey is ActionOnBreak.IgnoreBreak, the cancel key is ignored. If
'           BreakKey is ActionOnBreak.PromptUser, the user is given a ?Continue? message. If the
'           user selects "do not continue", the function returns ShellAndWaitResult.UserBreak = 6.
'           If the user selects "continue", the wait is continued.
'
'   Return values:
'            ShellAndWaitResult.Success = 0
'               indicates the the process completed successfully.
'            ShellAndWaitResult.Failure = 1
'               indicates that the Wait operation failed due to a Windows error.
'            ShellAndWaitResult.TimeOut = 2
'               indicates that the TimeOutMs interval timed out the Wait.
'            ShellAndWaitResult.InvalidParameter = 3
'               indicates that an invalid value was passed to the procedure.
'            ShellAndWaitResult.SysWaitAbandoned = 4
'               indicates that the system abandoned the wait.
'            ShellAndWaitResult.UserWaitAbandoned = 5
'               indicates that the user abandoned the wait via the cancel key (Ctrl+Break).
'               This happens only if BreakKey is set to ActionOnBreak.AbandonWait.
'            ShellAndWaitResult.UserBreak = 6
'               indicates that the user broke out of the wait after being prompted with
'               a ?Continue message. This happens only if BreakKey is set to
'               ActionOnBreak.PromptUser.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim TaskID As Long
Dim ProcHandle As Long
Dim WaitRes As Long
Dim Ms As Long
Dim MsgRes As VbMsgBoxResult
Dim SaveCancelKey As XlEnableCancelKey
Dim ElapsedTime As Long
Dim Quit As Boolean
Const ERR_BREAK_KEY = 18
Const DEFAULT_POLL_INTERVAL = 500

If Trim(ShellCommand) = vbNullString Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
End If

If TimeOutMs < 0 Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
ElseIf TimeOutMs = 0 Then
    Ms = WAIT_INFINITE
Else
    Ms = TimeOutMs
End If

Select Case BreakKey
    Case AbandonWait, IgnoreBreak, PromptUser
        ' valid
    Case Else
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
End Select

Select Case ShellWindowState
    Case vbHide, vbMaximizedFocus, vbMinimizedFocus, vbMinimizedNoFocus, vbNormalFocus, vbNormalNoFocus
        ' valid
    Case Else
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
End Select

On Error Resume Next
Err.Clear
TaskID = Shell(ShellCommand, ShellWindowState)
If (Err.Number <> 0) Or (TaskID = 0) Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
End If

ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskID)
If ProcHandle = 0 Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
End If

On Error GoTo ErrH:
SaveCancelKey = Application.EnableCancelKey
Application.EnableCancelKey = xlErrorHandler
WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
Do Until WaitRes = WAIT_OBJECT_0
    DoEvents
    Select Case WaitRes
        Case WAIT_ABANDONED
            ' Windows abandoned the wait
            ShellAndWait = ShellAndWaitResult.SysWaitAbandoned
            Exit Do
        Case WAIT_OBJECT_0
            ' Successful completion
            ShellAndWait = ShellAndWaitResult.Success
            Exit Do
        Case WAIT_FAILED
            ' attach failed
            ShellAndWait = ShellAndWaitResult.Failure
            Exit Do
        Case WAIT_TIMEOUT
            ' Wait timed out. Here, this time out is on DEFAULT_POLL_INTERVAL.
            ' See if ElapsedTime is greater than the user specified wait
            ' time out. If we have exceed that, get out with a TimeOut status.
            ' Otherwise, reissue as wait and continue.
            ElapsedTime = ElapsedTime + DEFAULT_POLL_INTERVAL
            If Ms > 0 Then
                ' user specified timeout
                If ElapsedTime > Ms Then
                    ShellAndWait = ShellAndWaitResult.TimeOut
                    Exit Do
                Else
                    ' user defined timeout has not expired.
                End If
            Else
                ' infinite wait -- do nothing
            End If
            ' reissue the Wait on ProcHandle
            WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
            
        Case Else
            ' unknown result, assume failure
            ShellAndWait = ShellAndWaitResult.Failure
            Exit Do
            Quit = True
    End Select
Loop

CloseHandle ProcHandle
Application.EnableCancelKey = SaveCancelKey
Exit Function

ErrH:
Debug.Print "ErrH: Cancel: " & Application.EnableCancelKey
If Err.Number = ERR_BREAK_KEY Then
    If BreakKey = ActionOnBreak.AbandonWait Then
        CloseHandle ProcHandle
        ShellAndWait = ShellAndWaitResult.UserWaitAbandoned
        Application.EnableCancelKey = SaveCancelKey
        Exit Function
    ElseIf BreakKey = ActionOnBreak.IgnoreBreak Then
        Err.Clear
        Resume
    ElseIf BreakKey = ActionOnBreak.PromptUser Then
        MsgRes = MsgBox("User Process Break." & vbCrLf & _
            "Continue to wait?", vbYesNo)
        If MsgRes = vbNo Then
            CloseHandle ProcHandle
            ShellAndWait = ShellAndWaitResult.UserBreak
            Application.EnableCancelKey = SaveCancelKey
        Else
            Err.Clear
            Resume Next
        End If
    Else
        CloseHandle ProcHandle
        Application.EnableCancelKey = SaveCancelKey
        ShellAndWait = ShellAndWaitResult.Failure
    End If
Else
    ' some other error. assume failure
    CloseHandle ProcHandle
    ShellAndWait = ShellAndWaitResult.Failure
End If

Application.EnableCancelKey = SaveCancelKey

End Function


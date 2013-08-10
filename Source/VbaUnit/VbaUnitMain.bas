Attribute VB_Name = "VbaUnitMain"
Option Explicit

Public Sub VbaUnitMain()
    
    Run '"AutoGenTester"
End Sub

Public Sub Run(Optional TestClassName As String)
    Dim r As TestRunner
    Set r = New TestRunner
    
    Dim hdebut As Single
    Dim hfin As Single
   
    hdebut = Timer 'debut
    Application.ScreenUpdating = False
'    Dim objShell As New Shell
'    objShell.MinimizeAll        ' Minimize all windows
'    Set objShell = Nothing

    r.Run TestClassName
    
    Application.ScreenUpdating = True
    hfin = Timer 'fin
    Debug.Print Format(hfin - hdebut, "Fixed"); "  second"
End Sub

Public Sub Prep(Optional ClassName As String)
    Dim AG As AutoGen
    Set AG = New AutoGen
    AG.Prep ClassName
End Sub

Public Function QW(s As String) As String
    QW = Chr(34) & s & Chr(34)
End Function

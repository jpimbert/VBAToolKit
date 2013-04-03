Attribute VB_Name = "VbaUnitMain"
Option Explicit

Public Sub VbaUnitMain()
    Run '"AutoGenTester"
End Sub

Public Sub Run(Optional TestClassName As String)
    Dim R As TestRunner
    Set R = New TestRunner
    R.Run TestClassName
End Sub

Public Sub Prep(Optional ClassName As String)
    Dim AG As AutoGen
    Set AG = New AutoGen
    AG.Prep ClassName
End Sub

Public Function QW(s As String) As String
    QW = Chr(34) & s & Chr(34)
End Function

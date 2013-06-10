VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITestManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Name, Run, Count
Public Property Let ClassName(RHS As String)
Public Property Get ClassName() As String
Public Function Run(Test As ITest, Optional RunManager As IRunManager) As IRunManager
Public Function CountTestCases(Test As ITest) As Long


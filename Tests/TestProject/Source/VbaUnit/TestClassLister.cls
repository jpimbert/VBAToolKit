VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestClassLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function TestClasses() As Collection
    Dim Components As VBComponents, Component As VBComponent
    Set TestClasses = New Collection
    Set Components = Application.VBE.ActiveVBProject.VBComponents
    For Each Component In Components
        If IsClassModule(Component.Type) And IsTestClassName(Component.name) Then
            TestClasses.Add Component.name
        End If
    Next
End Function

Public Function IsTestClassName(ComponentName As String) As Boolean
    If Len(ComponentName) <= 6 Then Exit Function
    IsTestClassName = Right(ComponentName, 6) Like "Tester"
End Function

Public Function IsClassModule(ComponentType As vbext_ComponentType) As Boolean
    IsClassModule = (ComponentType = vbext_ct_ClassModule)
End Function

Public Function NewTestClass(TestClassName As String) As ITest
    Dim test As ITest
    Set NewTestClass = SelectTestClass(TestClassName)
    If NewTestClass Is Nothing Then Exit Function
    NewTestClass.Manager.ClassName = TestClassName
End Function

Public Function SelectTestClass(TestClassName As String) As ITest
    Select Case TestClassName
        Case "fftAssertContentTester": Set SelectTestClass = New fftAssertContentTester
        Case "fftEnvironmentForTestTester": Set SelectTestClass = New fftEnvironmentForTestTester
        Case "fftMappingTester": Set SelectTestClass = New fftMappingTester
        Case "fftSQLStrategyTester": Set SelectTestClass = New fftSQLStrategyTester
        Case "fftSynchroWorkBookTester": Set SelectTestClass = New fftSynchroWorkBookTester
        Case "fftTableSheetTester": Set SelectTestClass = New fftTableSheetTester
        Case Else:
    End Select
End Function

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
        Case "vtkCreateProjectTester": Set SelectTestClass = New vtkCreateProjectTester
        Case "vtkPathUtilitiesTester": Set SelectTestClass = New vtkPathUtilitiesTester
        Case "vtkWorkbookIsOpenTester": Set SelectTestClass = New vtkWorkbookIsOpenTester
        Case "vtkProjectTester": Set SelectTestClass = New vtkProjectTester
        Case "vtkCreateTreeFolderTester": Set SelectTestClass = New vtkCreateTreeFolderTester
        Case "vtkImportExportUtilitiesTester": Set SelectTestClass = New vtkImportExportUtilitiesTester
        Case "vtkToolBarsTester": Set SelectTestClass = New vtkToolBarsTester
        Case "vtkExportOneModuleTester": Set SelectTestClass = New vtkExportOneModuleTester
        Case "vtkGitFunctionsTester": Set SelectTestClass = New vtkGitFunctionsTester
        Case "vtkExportConfigurationTester": Set SelectTestClass = New vtkExportConfigurationTester
        Case "vtkNormalizeTester": Set SelectTestClass = New vtkNormalizeTester
        Case "vtkTestUtilitiesTester": Set SelectTestClass = New vtkTestUtilitiesTester
        Case "vtkRecreateConfigurationTester": Set SelectTestClass = New vtkRecreateConfigurationTester
        Case "vtkConfigurationManagersTester": Set SelectTestClass = New vtkConfigurationManagersTester
        Case "vtkConfigurationManagerTester": Set SelectTestClass = New vtkConfigurationManagerTester
        Case "vtkFileSystemUtilitiesTester": Set SelectTestClass = New vtkFileSystemUtilitiesTester
        Case "vtkXMLUtilitiesTester": Set SelectTestClass = New vtkXMLUtilitiesTester
        Case "vtkModuleTester": Set SelectTestClass = New vtkModuleTester
        Case "vtkConfigurationTester": Set SelectTestClass = New vtkConfigurationTester
        Case "vtkXMLExportTester": Set SelectTestClass = New vtkXMLExportTester
        Case "vtkProjectsTester": Set SelectTestClass = New vtkProjectsTester
        Case "vtkReferenceTester": Set SelectTestClass = New vtkReferenceTester
        Case "vtkExcelUtilitiesTester": Set SelectTestClass = New vtkExcelUtilitiesTester
        Case Else:
    End Select
End Function

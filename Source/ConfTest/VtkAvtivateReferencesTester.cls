VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtkAvtivateReferencesTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Implements ITest
Implements ITestCase

Private mManager As TestCaseManager
Private mAssert As IAssert

Private Sub Class_Initialize()
    Set mManager = New TestCaseManager
End Sub

Private Property Get ITestCase_Manager() As TestCaseManager
    Set ITestCase_Manager = mManager
End Property

Private Property Get ITest_Manager() As ITestManager
    Set ITest_Manager = mManager
End Property

Private Sub ITestCase_SetUp(Assert As IAssert)
    Set mAssert = Assert
    

End Sub

Private Sub ITestCase_TearDown()
    'allow to iterate over all files in a directory
    Dir (vtkTestPath)
    'to not display messagebox how ask to save project
    Application.DisplayAlerts = False
    On Error Resume Next
    'Close Created WorkBook
    Workbooks(vtkTestProjectName & ".xlsm").Close
    'Delete Created WorkBook
    Kill vtkTestPath & "\" & vtkTestProjectName & "\Project\" & vtkTestProjectName & ".xlsm"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfProd"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\ConfTest"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source\VbaUnit"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Tests"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Source"
    RmDir vtkTestPath & "\" & vtkTestProjectName & "\Project"
    RmDir vtkTestPath & "\" & vtkTestProjectName
End Sub

Public Sub TestReferencesWasActivated()
 Dim x As Long
 Dim i As Integer
 
'this function will count extention number before executing function ,
'and extentions number must be equal to(initial extention number + added extention number )

   x = vtkCreateProject(vtkTestPath, vtkTestProjectName)
   i = Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count
   a = VtkActivateReferences(vtkTestProjectName & ".xlsm")

mAssert.Equals Workbooks(vtkTestProjectName & ".xlsm").VBProject.References.Count, i + 2, "the references number must be equal to" & i + 2
End Sub

Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestReferencesWasActivated"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "TestReferencesWasActivated": TestReferencesWasActivated
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub



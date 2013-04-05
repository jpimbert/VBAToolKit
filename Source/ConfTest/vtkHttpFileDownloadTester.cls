VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkHttpFileDownloadTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit


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

End Sub
'1st version
Public Sub Testinternetconnection()

Dim testconnection As vtktestconnection
Set testconnection = New vtktestconnection
mAssert.Should testconnection.Connexion = True, "The internet connection must be activated"

End Sub

Public Sub TestCreatefileDoesNothingWhenPathDoesntExist()
Dim InexistentPath As String
End Sub
 Public Sub TestCreateProjectDoesNothingWhenPathorsourceDoesntExist()
    ' The createProject must does nothing when path doesn't exist
    Dim InexistentPath As String
    Dim source As String
    
 Call télécharge_http(vtkTestPath:=InexistentPath, namefiletocopy:="tmp", sourcefilelink:=source)
    mAssert.Should vtkTestPath & namefiletocopy.txt = "", "The project main folder must exist"
    mAssert.Equals returnValue, 76, "The project creation must fail"
  


End Sub
Public Sub Testfilemustbecreated()
     
    Dim c As vtkhttpfiledownload    'instance de la classe vtkhttpfiledownload
    Set c = New vtkhttpfiledownload '
  
    Dim tmp1 As String
    Dim tmp2 As String
    tmp1 = vtkTestPath & "\"
    tmp2 = "test18h"

Call c.télécharge_http(tmp1, tmp2, vtkTestSource)

mAssert.Should vtkTestPath & "tmp" <> "", "The file must be created !! "

End Sub

Public Sub Testclosefilecreated()
 
End Sub
Public Sub testcloseinternetconnection()
 
End Sub


Private Function ITest_Suite() As TestSuite
    Set ITest_Suite = New TestSuite
    ITest_Suite.AddTest ITest_Manager.ClassName, "Testinternetconnection"
    ITest_Suite.AddTest ITest_Manager.ClassName, "TestCreatefileDoesNothingWhenPathDoesntExist"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Testfilemustbecreated"
    ITest_Suite.AddTest ITest_Manager.ClassName, "Testclosefilecreated"
End Function

Private Sub ITestCase_RunTest()
    Select Case mManager.methodName
        Case "Testinternetconnection": Testinternetconnection
        Case "TestCreatefileDoesNothingWhenPathDoesntExist": TestCreatefileDoesNothingWhenPathDoesntExist
        Case "Testfilemustbecreated": Testfilemustbecreated
        Case "Testclosefilecreated": Testclosefilecreated
        Case Else: mAssert.Should False, "Invalid test name: " & mManager.methodName
    End Select
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CountOfLines(module As CodeModule) As Long
    CountOfLines = module.CountOfLines
End Function

Public Function Lines(module As CodeModule, StartLine As Long, Count As Long) As String
    Lines = module.Lines(StartLine, Count)
End Function

Public Function SuiteMethodBody(TestMethodNames As Collection) As String
    Dim methodName As Variant
    SuiteMethodBody = SuiteHeader()
    For Each methodName In TestMethodNames
        SuiteMethodBody = SuiteMethodBody & vbCrLf & SuiteLine(CStr(methodName))
    Next
End Function

Public Function SuiteMethod() As String
    SuiteMethod = "Private Function ITest_Suite() As TestSuite"
End Function

Public Function SuiteHeader() As String
    SuiteHeader = "    Set ITest_Suite = New TestSuite"
End Function

Public Function SuiteLine(methodName As String) As String
    SuiteLine = "    ITest_Suite.AddTest ITest_Manager.ClassName, " & QW(methodName)
End Function

Public Function EndFunction() As String
    EndFunction = "End Function"
End Function

Public Function GetMethodBody(ClassName As String, methodName As String) As String
    Dim module As CodeModule, StartLine As Long, LineCount As Long, BodyLine As Long
    Set module = GetCodeModule(ClassName)
    If module Is Nothing Then Exit Function
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(module, methodName)
    LineCount = BodyLineCount(module, methodName)
    If StartLine = 0 Or LineCount = 0 Then Exit Function
    GetMethodBody = module.Lines(StartLine, LineCount)
'NO_METHOD:
'    Exit Function
End Function

Private Function BodyStartLine(module As CodeModule, methodName As String) As Long
    Dim BodyLine As Long, NumFuncLines As Long
    NumFuncLines = 1
    If module Is Nothing Then Exit Function
    On Error GoTo NO_METHOD
    BodyLine = module.ProcBodyLine(methodName, vbext_pk_Proc)
    BodyStartLine = BodyLine + NumFuncLines
NO_METHOD:
    Exit Function
End Function

Public Function BodyLineCount(module As CodeModule, methodName As String) As Long
    Dim StartLine As Long, LineCount As Long, BodyLine As Long, NumFuncLines As Long, NumEndFuncLines As Long, BStartLine As Long
    NumFuncLines = 1
    NumEndFuncLines = 1
    If module Is Nothing Then Exit Function
    On Error GoTo NO_METHOD
    StartLine = module.ProcStartLine(methodName, vbext_pk_Proc)
    LineCount = module.ProcCountLines(methodName, vbext_pk_Proc)
    Do While module.Lines(StartLine + LineCount - NumEndFuncLines, 1) = ""
        NumEndFuncLines = NumEndFuncLines + 1
    Loop
    BodyLine = module.ProcBodyLine(methodName, vbext_pk_Proc)
    BodyLineCount = LineCount - (BodyLine - StartLine) - NumFuncLines - NumEndFuncLines
    BStartLine = BodyStartLine(module, methodName)
NO_METHOD:
    Exit Function
End Function


Public Function GetCodeModule(ClassName) As CodeModule
    Dim Components As VBComponents, Component As VBComponent
    Set Components = Application.VBE.ActiveVBProject.VBComponents
    For Each Component In Components
        If Component.name = ClassName Then
            Set GetCodeModule = Component.CodeModule
            Exit Function
        End If
    Next
End Function

Public Sub ReplaceMethodBody(ClassName As String, methodName As String, NewMethodBody As String)
    Dim module As CodeModule, LineStart As Long, LineCount As Long, BodyLine As Long
    Set module = GetCodeModule(ClassName)
    If module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    DeleteMethodBody module, methodName
    InsertMethodBody module, methodName, NewMethodBody
NO_METHOD:
    Exit Sub
End Sub

Public Sub DeleteMethodBody(module As CodeModule, methodName As String)
    Dim StartLine As Long, LineCount As Long
    If module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(module, methodName)
    LineCount = BodyLineCount(module, methodName)
    If StartLine = 0 Or LineCount = 0 Then Exit Sub
    module.DeleteLines StartLine, LineCount
'NO_METHOD:
'    Exit Sub
End Sub

Private Sub InsertMethodBody(module As CodeModule, methodName As String, MethodBody As String)
    Dim StartLine As Long
    If module Is Nothing Then Exit Sub
    'On Error GoTo NO_METHOD
    StartLine = BodyStartLine(module, methodName)
    If StartLine = 0 Then Exit Sub
    module.InsertLines StartLine, MethodBody
'NO_METHOD:
'    Exit Sub
End Sub

Public Function GetTestMethods(ClassName As String) As Collection
    Dim module As CodeModule, LineNum As Long
    Set GetTestMethods = New Collection
    Set module = GetCodeModule(ClassName)
    If module Is Nothing Then Exit Function
    For LineNum = 1 To module.CountOfLines
        If IsTestMethodLine(module.Lines(LineNum, 1)) Then
            GetTestMethods.Add module.ProcOfLine(LineNum, vbext_pk_Set)
        End If
    Next
End Function

Private Function IsTestMethodLine(line As String) As Boolean
    IsTestMethodLine = Left(line, 15) Like "Public Sub Test"
End Function

Public Sub MakeSuite(ClassName As String)
    ReplaceMethodBody ClassName, "ITest_Suite", SuiteMethodBody(GetTestMethods(ClassName))
End Sub

Public Function RunTestLine(methodName As String) As String
    RunTestLine = "        Case " & QW(methodName) & ": " & methodName
End Function

Public Function RunTestHeader() As String
    RunTestHeader = "    Select Case mManager.MethodName"
End Function

Public Function RunTestFooter() As String
    RunTestFooter = "        Case Else: mAssert.Should False, " & QW("Invalid test name: ") & " & mManager.MethodName" & vbCrLf & _
                    "    End Select"
End Function

Public Function RunTestMethodBody(TestMethodNames As Collection) As String
    Dim methodName As Variant
    RunTestMethodBody = RunTestHeader()
    For Each methodName In TestMethodNames
        RunTestMethodBody = RunTestMethodBody & vbCrLf & RunTestLine(CStr(methodName))
    Next
    RunTestMethodBody = RunTestMethodBody & vbCrLf & RunTestFooter()
End Function

Public Sub MakeRunTest(ClassName As String)
    ReplaceMethodBody ClassName, "ITestCase_RunTest", RunTestMethodBody(GetTestMethods(ClassName))
End Sub

Public Sub Prep(Optional ClassName As String)
    Dim Classes As Collection, name As Variant
    MakeTestClassLister
    If ClassName = "" Then
        Set Classes = TestClasses()
    Else
        Set Classes = New Collection
        Classes.Add ClassName
    End If
    For Each name In Classes
        MakeSuite CStr(name)
        MakeRunTest CStr(name)
    Next
End Sub

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

Public Function TestClassHeader() As String
    TestClassHeader = "    Select Case TestClassName"
End Function

Public Function TestClassLine(TestClassName As String) As String
    TestClassLine = "        Case " & QW(TestClassName) & ": Set SelectTestClass = New " & TestClassName
End Function

Public Function TestClassFooter() As String
    TestClassFooter = "        Case Else:" & vbCrLf & _
                      "    End Select"
End Function

Public Function TestClassMethodBody(TestClassNames As Collection) As String
    Dim TestClassName As Variant
    TestClassMethodBody = TestClassHeader()
    For Each TestClassName In TestClassNames
        TestClassMethodBody = TestClassMethodBody & vbCrLf & TestClassLine(CStr(TestClassName))
    Next
    TestClassMethodBody = TestClassMethodBody & vbCrLf & TestClassFooter()
End Function

Public Sub MakeTestClassLister()
    ReplaceMethodBody "TestClassLister", "SelectTestClass", TestClassMethodBody(TestClasses())
End Sub

Attribute VB_Name = "vtkProjectCreationUtilities"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkInitializeVbaUnitNamesAndPathes
' Author    : Abdelfattah Lahbib
' Date      : 09/05/2013
' Purpose   : - Initialize DEV project ConfSheet with vbaunit module names and pathes
'             - Return True if module names and paths are initialized without error
'---------------------------------------------------------------------------------------
'
Public Function vtkInitializeVbaUnitNamesAndPathes(project As String) As Boolean
    Dim tableofvbaunitname(17) As String
        tableofvbaunitname(0) = "VbaUnitMain"
        tableofvbaunitname(1) = "Assert"
        tableofvbaunitname(2) = "AutoGen"
        tableofvbaunitname(3) = "IAssert"
        tableofvbaunitname(4) = "IResultUser"
        tableofvbaunitname(5) = "IRunManager"
        tableofvbaunitname(6) = "ITest"
        tableofvbaunitname(7) = "ITestCase"
        tableofvbaunitname(8) = "ITestManager"
        tableofvbaunitname(9) = "RunManager"
        tableofvbaunitname(10) = "TestCaseManager"
        tableofvbaunitname(11) = "TestClassLister"
        tableofvbaunitname(12) = "TesterTemplate"
        tableofvbaunitname(13) = "TestFailure"
        tableofvbaunitname(14) = "TestResult"
        tableofvbaunitname(15) = "TestRunner"
        tableofvbaunitname(16) = "TestSuite"
        tableofvbaunitname(17) = "TestSuiteManager"
    Dim i As Integer, cm As vtkConfigurationManager, ret As Boolean, nm As Integer, nc As Integer, ext As String
    Set cm = vtkConfigurationManagerForProject(project)
    nc = cm.getConfigurationNumber(vtkProjectForName(project).projectDEVName)
    ret = (nc > 0)
    For i = LBound(tableofvbaunitname) To UBound(tableofvbaunitname)
        nm = cm.addModule(tableofvbaunitname(i))
        ret = ret And (nm > 0)
        If i <= 0 Then      ' It's a Standard Module (WARNING, magical number)
            ext = ".bas"
           Else
            ext = ".cls"    ' It's a Class Module
        End If
        cm.setModulePathWithNumber path:="Source\VbaUnit\" & tableofvbaunitname(i) & ext, numModule:=nm, numConfiguration:=nc
    Next i
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function

'---------------------------------------------------------------------------------------
' Procedure : VtkAvtivateReferences
' Author    : Abdelfattah Lahbib
' Date      : 26/04/2013
' Purpose   : - check that workbook is open and activate VBIDE and +-scripting references
'---------------------------------------------------------------------------------------
Public Function VtkActivateReferences(workbookName As String)
    If VtkWorkbookIsOpen(workbookName) = True Then     'if the workbook is ope
        On Error Resume Next ' if the first extention is already activated, we will try to activate the second one
        Workbooks(workbookName).VBProject.References.AddFromGuid "{420B2830-E718-11CF-893D-00A0C9054228}", 0, 0  ' +- to activate Scripting : Microsoft scripting runtime
        Workbooks(workbookName).VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 0, 0 ' to activate VBIDE: Microsoft visual basic for applications extensibility 5.3
    End If
End Function

Attribute VB_Name = "vtkMainFunctions"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : vtkCreateProject
' Author    : JPI-Conseil
' Date      : 03/04/2013
' Purpose   : Create a tree folder for a new project
'               - Source containing ConfProd, ConfTest and VBAUnit
'               - Project containing the main Excel file for the project
'               - an empty Tests folder
'               - A Git repository is initialized for the project
'               - Create Xlsm Dev and Delivery project
'               - Rename 2 Project
'               - Activate missing References
'               - Initialize local repository for active project
' Return    : Long error number
'
'   L'extension "Microsoft Visual Basic For Application Extensibility" doit être activée
'
'   some problems : -you can't create 2 project with the same name
'
'---------------------------------------------------------------------------------------
'
Public Function vtkCreateProject(path As String, name As String, Optional displayError As Boolean = True) As Long
    
   
        'Application.DisplayAlerts = False 'to not display message that ask to save project
       On Error GoTo vtkCreateProject_Error

            ' Create the vtkProject attached to the new project
            Dim project As vtkProject
            Set project = vtkProjectForName(projectname:=name)
            
            ' Create main folder
            MkDir path & "\" & project.projectname
            ' Create Delivery folder
            MkDir path & "\" & project.projectname & "\" & "Delivery"
            ' Create Project folder
            MkDir path & "\" & project.projectname & "\" & "Project"
            ' Create Tests folder
            MkDir path & "\" & project.projectname & "\" & "Tests"
            ' Create GitLog Folder
            MkDir path & "\" & project.projectname & "\" & "GitLog"
            ' Create Source folder
            MkDir path & "\" & project.projectname & "\" & "Source"
            ' Create ConfProd folder
            MkDir path & "\" & project.projectname & "\" & "Source" & "\" & "ConfProd"
            ' Create ConfTest folder
            MkDir path & "\" & project.projectname & "\" & "Source" & "\" & "ConfTest"
            ' Create VbaUnit folder
            MkDir path & "\" & project.projectname & "\" & "Source" & "\" & "VbaUnit"
             
            'Save created project with xlsm extention
             Workbooks.Add.SaveAs (path & "\" & project.projectname & "\" & project.projectDEVStandardRelativePath), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(project.workbookDEVName).VBProject.name = project.projectDEVName
            'call function who activate references
            VtkActivateReferences (project.workbookDEVName)
            'initialize configuration Sheet with VBAUnit modules
            vtkInitializeVbaUnitNamesAndPathes project:=project.projectname
            ' Save Development Project Workbook
            Workbooks(project.workbookDEVName).Save
            
            'Create delivery workbook
            Workbooks.Add.SaveAs (path & "\" & name & "\" & project.projectStandardRelativePath), FileFormat:=(52) '52 is xlsm format
            'Rename Project
            Workbooks(project.workbookname).VBProject.name = project.projectname
            'call function who activate references
            VtkActivateReferences (project.workbookname)
            ' A module must be added in the Excel File for the project parameters to be saved
            Workbooks(project.workbookname).VBProject.VBComponents.Add ComponentType:=vbext_ct_StdModule
            ' Save and Close Delivery Project WorkBook
            Workbooks(project.workbookname).Close SaveChanges:=True
            
            Workbooks(project.workbookDEVName).Activate
            Dim GitRetval As String
            GitRetval = vtkInitializeGit()
    On Error GoTo 0
    vtkCreateProject = 0
    Exit Function
vtkCreateProject_Error:
    vtkCreateProject = Err.Number
If displayError Then MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vtkCreateProject of Module MainFunctions"
End Function

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
        
    Dim i As Integer, cm As vtkConfigurationManager, proj As vtkProject, ret As Boolean, nm As Integer, nc As Integer, ext As String
    
    Set cm = vtkConfigurationManagerForProject(project)
    
    Set proj = vtkProjectForName(projectname:=project)
    
    nc = cm.getConfigurationNumber(vtkProjectForName(project).projectDEVName)
    ret = (nc > 0)
    For i = LBound(tableofvbaunitname) To UBound(tableofvbaunitname)
        nm = cm.AddModule(tableofvbaunitname(i))
        ret = ret And (nm > 0)
        If i <= 0 Then      ' It's a Standard Module
            ext = ".bas"
           Else
            ext = ".cls"    ' It's a Class Module
        End If
        cm.setModulePathWithNumber path:="Source\VbaUnit\" & tableofvbaunitname(i) & ext, numModule:=nm, numConfiguration:=nc
      
        'export module from source workbook to the created project folder
        Workbooks(ThisWorkbook.name).VBProject.VBComponents(tableofvbaunitname(i)).Export (proj.ProjectFullPath & "\Source\VbaUnit\" & tableofvbaunitname(i) & ext)
        'import module from the new project folder to the new workbook
        Workbooks(proj.projectDEVName & ".xlsm").VBProject.VBComponents.Import (proj.ProjectFullPath & "\Source\VbaUnit\" & tableofvbaunitname(i) & ext)
    Next i
    vtkInitializeVbaUnitNamesAndPathes = ret
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkAddModule
' Author    : user
' Date      : 20/06/2013
' Purpose   : -add a new module to a workbook
'             - take modulename and extention on parameter ( form ,std module or classe)
'---------------------------------------------------------------------------------------
'
Public Function VtkAddOneModule(moduleName As String, ext As Integer) As String
    
    Dim cm As vtkConfigurationManager
    Dim RetValCreateStandarMod As String
    Dim proj As vtkProject
    Dim nm As Integer
    Dim nc As Integer
    Dim ret As Boolean
    Dim fso As New FileSystemObject
    Dim ProjectPath As String
    Dim projname As String
    
    projname = Left(ActiveWorkbook.VBProject.name, (Len(ActiveWorkbook.VBProject.name) - 4))

    Set cm = vtkConfigurationManagerForProject(projname)
    Set proj = vtkProjectForName(projectname:=projname)
   ' Debug.Print proj.projectname
    
    nm = cm.AddModule(moduleName)
    nc = cm.getConfigurationNumber(vtkProjectForName(projname).projectDEVName)
    ProjectPath = fso.GetParentFolderName(ActiveWorkbook.path)

    ret = (nc > 0)
      
      Select Case ext
        
        Case 1 '1module : export to confprod
          
          cm.setModulePathWithNumber path:="\Source\ConfProd\" & moduleName & ".BAS", numModule:=nm, numConfiguration:=nc
          CreateAndExportNewModule ActiveWorkbook, 1, moduleName, ProjectPath & "\Source\ConfProd\" & moduleName & ".BAS"
       Case 2 '2 class module : export to ConfTest or ConfProd
         
            If Right(moduleName, 6) Like "Tester" Then ' verify if modulename end is like Tester
     '
                ' This Document is a test module export to confTest
          cm.setModulePathWithNumber path:="\Source\ConfTest\" & moduleName & ".CLS", numModule:=nm, numConfiguration:=nc
          CreateAndExportNewModule ActiveWorkbook, 2, moduleName, ProjectPath & "\Source\ConfTest\" & moduleName & ".CLS"
               Else
    
                'the document is a classmodule export to confprod
         cm.setModulePathWithNumber path:="\Source\ConfProd\" & moduleName & ".CLS", numModule:=nm, numConfiguration:=nc
         CreateAndExportNewModule ActiveWorkbook, 2, moduleName, ProjectPath & "\Source\ConfProd\" & moduleName & ".CLS"
            End If
        Case 3 '3 forms
        
       '         'the document is a classmodule export to confprod

        cm.setModulePathWithNumber path:="\Source\ConfProd\" & moduleName & ".FRM", numModule:=nm, numConfiguration:=nc
        CreateAndExportNewModule ActiveWorkbook, 3, moduleName, ProjectPath & "\Source\ConfProd\" & moduleName & ".FRM"
   '    Exit Function
          
       End Select

     
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateAndImportModule
' Author    : user
' Date      : 21/06/2013
' Purpose   : - create and export modules
'---------------------------------------------------------------------------------------
'

Sub CreateAndExportNewModule(ByVal wb As Workbook, _
    ByVal ModuleTypeIndex As Integer, ByVal NewModuleName As String, ByVal path)
 
    Dim VBC As VBComponent
    Set VBC = Nothing
    Set VBC = wb.VBProject.VBComponents.Add(ModuleTypeIndex)
    VBC.name = NewModuleName
    VBC.Export path
End Sub


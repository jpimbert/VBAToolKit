Attribute VB_Name = "vtktmptest"
'Option Explicit

Public Function testfn()
a = vtkCreateProject(vtkTestPath, "Lah")

C = vtkExportAll("VBAToolKit.xlsm")
f = vtkImportModuleToDevWorkbook()
Debug.Print f
End Function


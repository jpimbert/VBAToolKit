VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : ThisWorkbook
' Author    : Jean-Pierre Imbert
' Date      : 09/08/2013
' Purpose   : Manage the toolbars during install/uninstall of addin
'
' WARNING :
'      This module exists in slightly different Prod and Test versions
'          - explicit call to VBAToolKit project for Prod version
'          - explicit call to VBAToolKit_DEV project for Test version
' TO DO
'   Dans la version PROD, recréer la barre VBE à l'ouverture du fichier (la VBE est temporaire)
'   Dans la version DEV, exporter les modules modifiés lors de la sauvegarde
'---------------------------------------------------------------------------------------

Option Explicit

Private Sub Workbook_AddinInstall()
    VBAToolKit_DEV.vtkToolBars.vtkCreateToolbars
End Sub

Private Sub Workbook_AddinUninstall()
    VBAToolKit_DEV.vtkToolBars.vtkDeleteToolbars
End Sub

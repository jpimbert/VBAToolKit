Attribute VB_Name = "EssaisDivers"
Rem
Rem Configuration requise
Rem     VB extensibility
Rem     ADO
Rem
Rem Programme pour la suite
Rem
Rem - Prendre en compte les trois types de table
Rem     - Table principale (mainTable)
Rem         avec un recordKey (local à une base) et un refKey (unique sur plusieurs bases)
Rem         Les espaces de refKey sont partitionnés sur les deux bases à synchroniser
Rem         La synchronisation se fait de la base réduite (un seul espace de refKey) vers la
Rem        base complète qui regroupe tous les refKeys
Rem         Le recordKey doit être remplacé par refKey lors de l'export vers Excel (dans toutes les tables)
Rem         Puis le refKey doit être retraduit en recordKey local lors de l'import dans la base complète
Rem     - Table secondaire (secondaryTable)
Rem         informations multiples (relation 1-n) liées à mainTable par le recordKey
Rem         Il faut utiliser le rowguid pour détecter les modifications (créé côté base réduite)
Rem     - Table accessoire (accessoryTable), items identifiés par leur propre recordkey
Rem         relations n-n liées à mainTable par les recordKey par une table (relTable)
Rem         Il faut utiliser le rowguid pour détecter les modifications (créé côté base complète)
Rem         Les tables accessoires se synchronisent de la base complète vers la base réduite
Rem
Rem - réaliser le wiki
Rem - mettre sou GitHub
Rem - Créer l'environnement de projet (chargement des modules)
Rem

Option Explicit

'---------------------------------------------------------------------------------------
' Module    : EssaisDivers
' Author    : Jean-Pierre Imbert
' Date      : 24/11/2012
' Purpose   : Démonstration de synchronisation de bases de données
'---------------------------------------------------------------------------------------

Public Sub TestProgressViewWithBigExport()
    Dim mSavedActivatedWorkBook As Workbook
    Dim mSynchroWorkBook As New fftSynchroWorkbook
    Dim env As New fftEnvironmentForTestAccess2007, ws As Worksheet
    
    Set mSavedActivatedWorkBook = Application.ActiveWorkbook
    Set mSynchroWorkBook.environment = env
   
    env.destinationDatabase = fftTestDatabaseCompleteEmpty
    env.originDatabase = fftTestDatabaseReducedNoBase
    env.testWorkBookPolicy = fftTestWorkBookCompleteBig1       ' Very big export
    
    env.testWorkBook.Activate
    Call PopulateSynchroWorkBookForTest(mSynchroWorkBook)
    
    Rem
    Rem     Test avec les quantités suivantes
    Rem         environ 10000 FFTs, 2000 commentaires, 2000 documents et 5000 relations FFT-Docs
'    Call populateTestDatabase(mSynchroWorkBook.environment.databaseConnection(fftReducedDatabase), 10000, 2000, 2000, 5000)
    Call populateTestDatabase(mSynchroWorkBook.environment.databaseConnection(fftCompleteDatabase), 15000, 3000, 2000, 5000)
    Rem
    Rem     Les résultats sont les suivants :
    Rem         - import quasi immédiat
    Rem         - export environ 1 mn, inchangé avec une transaction pour l'export
    Rem

'    Call populateTestDatabase(mSynchroWorkbook.environment.databaseConnection(fftReducedDatabase), 100, 20, 20, 50)
'    Call populateTestDatabase(mSynchroWorkBook.environment.databaseConnection(fftCompleteDatabase), 150, 30, 20, 50)

    Debug.Print Now()
    Call mSynchroWorkBook.Synchronize
    Debug.Print Now()
    
'    mSynchroWorkBook.Workbook.Close (True)
    mSynchroWorkBook.environment.databaseClose (fftReducedDatabase)
    mSynchroWorkBook.environment.databaseClose (fftCompleteDatabase)
    Set mSynchroWorkBook = Nothing
    mSavedActivatedWorkBook.Activate
    Set mSavedActivatedWorkBook = Nothing
   
End Sub

Public Sub testProgressForm()
    Dim n As Integer, progressView As IfftProgressView
    n = 10
    
    Set progressView = fftProgressForm
    progressView.text = "Avancement " & n & "%"
    progressView.percentage = n
    fftProgressForm.Show
End Sub

'---------------------------------------------------------------------------------------
' Trouvé sur Internet pour mettre à jour une base SQL-Server à partir d'un range Excel
Sub ExceltoSQLUpload(gc_strServerAddress As String, gc_strDatabase As String, strTableName As String)
    Dim Cnn             As Object
    Dim wbkOpen         As Workbook
    Dim fd                          As FileDialog
    Dim objfl                       As Variant
    Dim rngName                     As Range
     
    Set fd = Application.FileDialog(msoFileDialogOpen)
    With fd
        .ButtonName = "Open"
        .AllowMultiSelect = False
        .Filters.Add "Text Files", "*.xlsm,*.xlsx;*.xls", 1
        .Title = "Select Raw Data File...."
        .InitialView = msoFileDialogViewThumbnail
        If .Show = -1 Then
            For Each objfl In .SelectedItems
                .Execute
            Next objfl
        End If
        On Error GoTo 0
    End With
    Set wbkOpen = ActiveWorkbook
    Set fd = Nothing
    Set rngName = Application.InputBox("Select Range to Upload in newly opended file", , , , , , , 8)
    rngName.name = "TempRange"
    strFileName = wbkOpen.FullName
     
    Set Cnn = CreateObject("ADODB.Connection")
    Cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes"";"
     
    nSQL = "INSERT INTO [odbc;Driver={SQL Server};" & _
    "Server=" & gc_strServerAddress & ";Database=" & gc_strDatabase & "]." & strTableName
    nJOIN = " SELECT * from [TempRange]"
    Cnn.Execute nSQL & nJOIN
    MsgBox "Uploaded Successfully", vbInformation, "Say Thank you to me"
     
    wbkOpen.Close
    Set wbkOpen = Nothing
End Sub


Public Sub TestGetGUID()
    MsgBox GetGUID, vbInformation, "GUID Generated"
End Sub
 
Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function


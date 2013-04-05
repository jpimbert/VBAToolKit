VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkhttpfiledownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtkhttpfiledownload
' Author    : user
' Date      : 05/04/2013
' Purpose   :-PATH DESTINATION = PROJECT\Source\VBAUNIT
'              - namefiletocopy name of file to read write
'---------------------------------------------------------------------------------------

'Option Explicit

 Private Declare Function OuvreInternet Lib "wininet" _
    Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, _
    ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function fermeInternet Lib "wininet" _
    Alias "InternetCloseHandle" (ByVal hInet As Long) As Integer
    
Private Declare Function code_page Lib "wininet" _
    Alias "InternetReadFile" (ByVal hFile As Long, ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    
 Private Declare Function Ouvrepage Lib "wininet" _
    Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwflags As Long, _
    ByVal dwContext As Long) As Long

Sub télécharge_http(pathdestination As String, namefiletocopy As String, sourcefilelink As String)
Dim texte_code As String * 1024

fich = InputBox("adresse Internet du fichier à télécharger ?", _
    "téléchargement HTTP", sourcefilelink)

'recherche extension du fichier
extn = ""
txt = fich
encor:
extn = Right(txt, 1) & extn
txt = Left(txt, Len(txt) - 1)
If Left(extn, 1) <> "." Then GoTo encor
If extn = ".fr" Or extn = ".com" Then extn = ".html"

'connection au fichier à télécharger
internet = OuvreInternet("toto", 1, vbNullString, vbNullString, 0) 'ouvre Internet
URL = Ouvrepage(internet, fich, vbNullString, _
    ByVal 0&, &H80000000, ByVal 0&) 'ouvre la page Web

'lecture du fichier par paquet de 1024 bytes
txt = ""
nb_caractères_lus = 1
Do While nb_caractères_lus > 0
code_page URL, texte_code, 1024, nb_caractères_lus
txt = txt & Left(texte_code, nb_caractères_lus)
Loop

'ménage
fermeInternet URL 'ferme la page
fermeInternet internet 'ferme Internet

'recopie dans un fichier
Open pathdestination & namefiletocopy & extn For Output As #1
Print #1, txt
Close #1

'ouverture du fichier téléchargé
'ThisWorkbook.FollowHyperlink "C:\Users\user\Desktop\VBATKIT\rien" & extn, , True

End Sub




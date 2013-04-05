VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtktestconnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : vtktestconnection
' Author    : user
' Date      : 05/04/2013
' Purpose   :- return true ,if internet connection work
'
'---------------------------------------------------------------------------------------


'Option Explicit

 Function Connexion()
Set WshShell = CreateObject("WScript.Shell")
PINGFlag = Not CBool(WshShell.Run("ping -n 1 www.google.com", 0, True))
 If PINGFlag = True Then
 'Successful ping
 Connexion = True
 Else
 'Unsuccessful ping
 Connexion = False

 End If
  
End Function


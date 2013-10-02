VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Should(Condition As Boolean, Optional Comment As String)
Public Sub Equals(actual As Variant, expected As Variant, Optional Comment As String)
Public Sub Delta(actual As Variant, expected As Variant, Delta As Variant, Optional Comment As String)

Attribute VB_Name = "vtkTypes"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkTypes
' Author    : Jean-Pierre Imbert
' Date      : 10/06/2014
' Purpose   :
'
' Copyright 2014 Skwal-Soft (http://skwalsoft.com)
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'---------------------------------------------------------------------------------------

'   Combo type instead of defining a vtkReference class
Public Type vtkReference
    name As String  ' Mandatory for each reference
    guid As String  ' guid and path are exclusive each other
    path As String
End Type



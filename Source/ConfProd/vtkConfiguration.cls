VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConfiguration
' Author    : Jean-Pierre Imbert
' Date      : 28/08/2013
' Purpose   : Contains informations relative to a VTK Configuration
'               - 'name' of the configuration (wrapper to 'parent' methods)
'               - 'parent' is the configuration manager in which is included the configuration
'               - 'number' (read only) is the number of the configuration in the configuration manager
'               - 'path' is the full path for the Excel workbook containing the configuration (wrapper to 'parent' methods)
'               - 'modules' is a collection of vtkModule objects, all the modules contained in the configuration with non null path (wrapper to 'parent' methods)
'
' Copyright 2013 Skwal-Soft (http://skwalsoft.com)
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

Private m_parent As vtkConfigurationManager             ' A configuration is attached to a Configuration Manager
Private m_number As Integer


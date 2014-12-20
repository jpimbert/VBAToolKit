Attribute VB_Name = "vtkConstants"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkConstants
' Author    : Jean-Pierre Imbert
' Date      : 21/08/2013
' Purpose   :
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

Public Const VTK_OK = 0
Public Const VTK_UNEXPECTED_ERROR = 2001
Public Const VTK_WRONG_FILE_PATH = 2075
Public Const VTK_WRONG_FOLDER_PATH = 2076
Public Const VTK_FORBIDDEN_PARAMETER = 2077
Public Const VTK_WRONG_GENERIC_EXTENSION = 2078

Public Const VTK_GIT_NOT_INSTALLED = 3000
Public Const VTK_GIT_ALREADY_INITIALIZED_IN_FOLDER = 3001
Public Const VTK_GIT_PROBLEM_DURING_INITIALIZATION = 3003

Public Const VTK_MODULE_NOTATTACHED = 4001              ' The module must be attached to a configuration
Public Const VTK_INEXISTANT_CONFIGURATION = 4002        ' Unknown configuration
Public Const VTK_WORKBOOK_NOTOPEN = 4003
Public Const VTK_WORKBOOK_ALREADY_OPEN = 4004
Public Const VTK_NO_SOURCE_FILES = 4005
Public Const VTK_OBSOLETE_CONFIGURATION_SHEET = 4006
Public Const VTK_NOTINITIALIZED = 4007
Public Const VTK_ALREADY_INITIALIZED = 4008
Public Const VTK_INVALID_FIELD = 4009
Public Const VTK_TEMPLATE_NOT_FOUND = 4010
Public Const VTK_INVALID_XML_FILE = 4011
Public Const VTK_READONLY_FILE = 4012

Public Const VTK_UNEXPECTED_CHAR = 5001
Public Const VTK_UNEXPECTED_EOS = 5002

Public Const VTK_FILE_NOT_FOUND = 5003
Public Const VTK_DOESNT_COPY_FOLDER = 5004
Public Const VTK_FOLDER_NOT_FOUND = 5005
Public Const VTK_FILE_OPEN_OR_LOCKED = 5006

Public Const VTK_WORKBOOK_NOT_OPEN = 6001
Public Const VTK_WORKBOOK_NOT_INITIALIZED = 6002
Public Const VTK_PROJECT_NOT_INITIALIZED = 6003
Public Const VTK_DOM_NOT_INITIALIZED = 6004

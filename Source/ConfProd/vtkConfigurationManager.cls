VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vtkConfigurationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------------------
' Interface : vtkConfigurationManager
' Author    : Jean-Pierre Imbert
' Date      : 06/07/2014
' Purpose   : Interface for Configurations, References and Modules management of a VTK Project
'               - a Configuration is an Excel file containing Modules
'               - for each configuration, a Module can be imported/exported to a specific path relative to the VTK Project directory
'               - each configuration is a VBA project included in a specific Excel file whose path is relative to the VTK Project directory
'               - for each configuration, the set of VBA references is managed
'               - some other parameters are managed for each configration
' NOTE      : This interface is adopted by Excel and XML versions of a configuration manager
'               - vtkConfigurationManagerExcel and vtkConfigurationManagerXML
'             Only the getter methods are concerned by the interface, for now only the Excel version
'             contains setter methods for the configuration parameters
'
' Usage:
'   - Each instance of Configuration Manager is attached to a VTK project)
'       - the method vtkConfigurationManagerForProject of module vtkConfigurationManagers give the instance attached to a project, or create it
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

' init (avec XMLPath spécifique pour XML, avant d'avoir la gestion centralisée des projets)

' rootPath
' projectName (uniquement en Get pour la première version avant gestion des projets)

' configurationCount
' configurations
' configuration
' getConfigurationNumber
' getConfigurationPath
' getConfigurationPathWithNumber
' setConfigurationPathWithNumber <- (dans vtkConfiguration, ne sont utilisés que pour les tests)
' getConfigurationProjectNameWithNumber
' setConfigurationProjectNameWithNumber <-
' getConfigurationCommentWithNumber
' setConfigurationCommentWithNumber <-
' getConfigurationTemplateWithNumber
' setConfigurationTemplateWithNumber <-
' --> Soit on supprime les setters inutiles, soit on les garde dans l'interface, avec une erreur dans XML

' moduleCount
' modules
' module
' getModulePathWithNumber
' setModulePathWithNumber <- (dans vtkModule, non utilisé à part dans les tests)

' references
' getConfigurationReferencesWithNumber

' Stratégie de refactorisation
'  - implémentation des méthodes d'interface au fur et à mesure
'  - implémentation des tests en utilisant l'IF à la place du Excel au fur et à mesure

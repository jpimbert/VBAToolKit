Attribute VB_Name = "vtkXMLFilesCreator"
'---------------------------------------------------------------------------------------
' Module    : vtkXMLFilesCreator
' Author    : Lucas Vitorino
' Purpose   : Contain the utilities used to create XML files :
'               - vtkConfigurations sheets
'               - rememberedProjects sheet
'               - DTDs
'             These utilities are mostly untested.
'
' Usage:
'   - Each instance of Configuration Manager is attached to the DEV Excel Workbook of a project
'       - the method vtkConfigurationManagerForProject give the instance attached to a workbook, or create it
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

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : createProjectXMLSheet
' Author    : Lucas Vitorino
' Purpose   : Create the XML sheet of a fully initialized project at the given path :
'                   - Delivery and DEV configuration
'                   - VBAUnit modules
'                   - references (not implemented yet)
' Notes     : - By default he XML file will look for the DTD "vtkConfigurationsDTD.dtd" located in the Template folder
'               of the project.
'             - The folder structure is supposed to be a standard VBAToolKit project structure.
'---------------------------------------------------------------------------------------
'
Public Sub createInitializedXMLSheetForProject(sheetPath As String, _
                                    projectName As String, _
                                    Optional dtdPath As String = "../Templates/vtkConfigurationsDTD.dtd", _
                                    Optional addReferenceToSelf As Boolean = True)

    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=sheetpath, Overwrite:=True)
  
    Dim deliveryConfId As String
    Dim devConfId As String
    deliveryConfId = "c01"
    devConfId = "c02"
  
    ' Create the XML preamble
    xmlFile.WriteLine Text:="<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>"
    xmlFile.WriteLine Text:="<!DOCTYPE vtkConf SYSTEM """ & dtdPath & """>"
    xmlFile.WriteLine Text:="<vtkConf>"
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create the info object
    xmlFile.WriteLine Text:="    <info>"
    xmlFile.WriteLine Text:="        <vtkConfigurationsVersion>1.0</vtkConfigurationsVersion>"
    xmlFile.WriteLine Text:="        <projectName>" & projectName & "</projectName>"
    xmlFile.WriteLine Text:="    </info>"
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create the 2 configurations
    xmlFile.WriteLine Text:="    <configuration cID=""" & deliveryConfId & """>"
    xmlFile.WriteLine Text:="        <name>" & projectName & "</name>"
    xmlFile.WriteLine Text:="        <path>Delivery\" & projectName & ".xlsm</path>"
    xmlFile.WriteLine Text:="    </configuration>"
    xmlFile.WriteLine Text:="    <configuration cID=""" & devConfId & """>"
    xmlFile.WriteLine Text:="        <name>" & projectName & "_DEV</name>"
    xmlFile.WriteLine Text:="        <path>Project\" & projectName & "_DEV.xlsm</path>"
    xmlFile.WriteLine Text:="    </configuration>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create all the VBAUnit modules
    Dim i As Integer
    Dim moduleName As String
    Dim module As VBComponent
    For i = 1 To vtkVBAUnitModulesList.Count
        moduleName = vtkVBAUnitModulesList.Item(i)
        Set module = ThisWorkbook.VBProject.VBComponents(moduleName)
        xmlFile.WriteLine Text:="    <module mID=""m" & i & """>"
        xmlFile.WriteLine Text:="        <name>" & moduleName & "</name>"
        xmlFile.WriteLine Text:="        <modulePath confId=""" & devConfId & """>" & vtkStandardPathForModule(module) & "</modulePath>"
        xmlFile.WriteLine Text:="    </module>"
        xmlFile.WriteBlankLines Lines:=1
    Next i

    ' Create the references for the DEV and the delivery configurations
    Dim tmpRef As VBIDE.Reference
    For Each tmpRef In listOfDefaultReferences
        xmlFile.WriteLine Text:="    <reference confIDs=""" & deliveryConfId & " " & devConfId & """>"
        xmlFile.WriteLine Text:="        <name>" & tmpRef.name & "</name>"
        xmlFile.WriteLine Text:="        <guid>" & tmpRef.guid & "</guid>"
        xmlFile.WriteLine Text:="    </reference>"
        xmlFile.WriteBlankLines Lines:=1
    Next
    
    ' Create the reference "to self" only for the DEV configuration
    If addReferenceToSelf Then
        xmlFile.WriteLine Text:="    <reference confIDs=""" & devConfId & """>"
        xmlFile.WriteLine Text:="        <name>" & ThisWorkbook.VBProject.name & "</name>"
        xmlFile.WriteLine Text:="        <path>" & ThisWorkbook.FullName & "</path>"
        xmlFile.WriteLine Text:="    </reference>"
        xmlFile.WriteBlankLines Lines:=1
    End If

    ' Close the file
    xmlFile.WriteLine Text:="</vtkConf>"
    xmlFile.Close

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createDTDForVtkConfigurations
' Author    : Lucas Vitorino
' Purpose   : Create a DTD sheet for vtkConfigurations, that is to say the XML sheet describing
'             a project with all its configurations.
'---------------------------------------------------------------------------------------
'
Public Sub createDTDForVtkConfigurations(sheetPath As String)
    
    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=sheetpath, Overwrite:=True)
    
    xmlFile.WriteLine Text:="<!ELEMENT vtkConf (info,configuration*, module*, reference*)>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="   <!ELEMENT info (vtkConfigurationsVersion,projectName)>"
    xmlFile.WriteLine Text:="       <!ELEMENT vtkConfigurationsVersion (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ELEMENT projectName (#PCDATA)>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="   <!ELEMENT configuration (name,path,templatePath?,title?,comment?)>"
    xmlFile.WriteLine Text:="    <!ATTLIST configuration cID ID #REQUIRED>"
    xmlFile.WriteLine Text:="       <!ELEMENT name (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ELEMENT path (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ELEMENT templatePath (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ELEMENT title (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ELEMENT comment (#PCDATA)>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="   <!ELEMENT module (name, modulePath*)>"
    xmlFile.WriteLine Text:="    <!ATTLIST module mID ID #REQUIRED>"
    xmlFile.WriteLine Text:="       <!ELEMENT modulePath (#PCDATA)>"
    xmlFile.WriteLine Text:="       <!ATTLIST modulePath"
    xmlFile.WriteLine Text:="           confId IDREF #REQUIRED"
    xmlFile.WriteLine Text:="       >"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="   <!ELEMENT reference (name, (guid|path))>"
    xmlFile.WriteLine Text:="    <!ATTLIST reference confIDs IDREFS #REQUIRED>"
    xmlFile.WriteLine Text:="       <!ELEMENT guid (#PCDATA)>"
  
    xmlFile.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : createRememberedProjectsXMLSheet
' Author    : Lucas Vitorino
' Purpose   : Create a blank "remembered projects" xml sheet at the given path.
' Notes     : - By default he XML file will look for the DTD "vtkRememberedProjects.dtd" located in the Template folder
'               of the project.
'             - The folder structure is supposed to be a standard VBAToolKit project structure.
'---------------------------------------------------------------------------------------
'
Public Sub createRememberedProjectsXMLSheet(sheetPath As String, _
                                            Optional dtdPath As String = "../Templates/vtkRememberedProjectsDTD.dtd")

    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=sheetpath, Overwrite:=True)

    xmlFile.WriteLine Text:="<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>"
    xmlFile.WriteLine Text:="<!DOCTYPE vtkConf SYSTEM """ & dtdPath & """>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="<rememberedProjects>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="    <info>"
    xmlFile.WriteLine Text:="        <version>1.0</version>"
    xmlFile.WriteLine Text:="    </info>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="</rememberedProjects>"
    
    xmlFile.Close
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : createDTDforRememberedProjects
' Author    : Lucas Vitorino
' Purpose   : Create a DTD sheet for rememberedProjects, that is to say the XML sheet describing
'             the projects rememberd by VBAToolKit (name, rootFolder, xmlRelativePath)
'---------------------------------------------------------------------------------------
'
Public Sub createDTDforRememberedProjects(sheetPath As String)
    
    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=sheetpath, Overwrite:=True)
    
    xmlFile.WriteLine Text:="<!ELEMENT rememberedProjects (info,project*)>"
    xmlFile.WriteLine Text:="    <!ELEMENT info (version)>"
    xmlFile.WriteLine Text:="        <!ELEMENT version (#PCDATA)>"
    xmlFile.WriteLine Text:="    <!ELEMENT project (name,rootFolder,xmlRelativePath)>"
    xmlFile.WriteLine Text:="        <!ELEMENT name (#PCDATA)>"
    xmlFile.WriteLine Text:="        <!ELEMENT rootFolder (#PCDATA)>"
    xmlFile.WriteLine Text:="        <!ELEMENT xmlRelativePath (#PCDATA)>"
    
    xmlFile.Close
    
End Sub

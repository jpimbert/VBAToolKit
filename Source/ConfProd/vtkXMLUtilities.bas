Attribute VB_Name = "vtkXMLUtilities"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : vtkXMLUtilities
' Author    : Lucas Vitorino
' Purpose   : Provide utilities to support XML.
'
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

'---------------------------------------------------------------------------------------
' Procedure : escapedString
' Author    : Jean-Pierre IMBERT
' Date      : 19/12/2014
' Purpose   : escape < and > characters for XML PCDATA
'---------------------------------------------------------------------------------------
'
Public Function escapedString(str As String) As String
    Dim str1 As String
    str1 = Replace(str, "<", "&lt;")
    escapedString = Replace(str1, ">", "&gt;")
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkExportConfigurationsAsXML
' Author    : Jean-Pierre IMBERT
' Date      : 07/11/2013
' Purpose   : Export the configurations of a project as a new XML file
'             - the "new" XML file format is the one designed for configuration management as XML
'             - this subroutine is temporary, dedicated to prepare the migration from
'               Excel sheet management of configurations to XML file management
' Raises    : - VTK_WORKBOOK_NOT_OPEN if the _DEV workbook containing the configuration sheet is not opened
'             - VTK_WRONG_FILE_PATH if the file path couldn't be created
'---------------------------------------------------------------------------------------
'
Public Sub vtkExportConfigurationsAsXML(projectName As String, filePath As String)

   On Error GoTo vtkExportConfigurationsAsXML_Error

    ' Get the configurationManager of the project to export
    Dim cm As vtkConfigurationManager
    Set cm = vtkConfigurationManagerForProject(projectName)
    If cm Is Nothing Then
        Err.Raise Number:=VTK_WORKBOOK_NOT_OPEN
    End If
    
    ' Create a new XML configuration file
    Dim fso As New FileSystemObject
    Dim xmlFile As TextStream
    Set xmlFile = fso.CreateTextFile(fileName:=filePath, Overwrite:=True)

    ' Create the XML preamble
    xmlFile.WriteLine Text:="<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""yes""?>"
    xmlFile.WriteLine Text:="<!DOCTYPE vtkConf ["
    xmlFile.WriteLine Text:="    <!ELEMENT vtkConf (info, reference*, configuration*, module*)>"
    xmlFile.WriteLine Text:="        <!ELEMENT info (vtkConfigurationsVersion,projectName)>"
    xmlFile.WriteLine Text:="                <!ELEMENT vtkConfigurationsVersion (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT projectName (#PCDATA)>"
    xmlFile.WriteLine Text:="        <!ELEMENT configuration (name,path,templatePath?,title?,comment?,password?)>"
    xmlFile.WriteLine Text:="         <!ATTLIST configuration cID ID #REQUIRED>"
    xmlFile.WriteLine Text:="         <!ATTLIST configuration refIDs IDREFS #IMPLIED>"
    xmlFile.WriteLine Text:="                <!ELEMENT name (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT path (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT templatePath (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT title        (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT comment (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ELEMENT password (#PCDATA)>"
    xmlFile.WriteLine Text:="        <!ELEMENT module (name, modulePath*)>"
    xmlFile.WriteLine Text:="         <!ATTLIST module mID ID #REQUIRED>"
    xmlFile.WriteLine Text:="                <!ELEMENT modulePath (#PCDATA)>"
    xmlFile.WriteLine Text:="                <!ATTLIST modulePath confId IDREF #REQUIRED>"
    xmlFile.WriteLine Text:="        <!ELEMENT reference (name, (guid|path))>"
    xmlFile.WriteLine Text:="         <!ATTLIST reference refID ID #REQUIRED>"
    xmlFile.WriteLine Text:="                <!ELEMENT guid (#PCDATA)>"
    xmlFile.WriteLine Text:="]>"
    
    xmlFile.WriteLine Text:="<vtkConf>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="    <info>"
    xmlFile.WriteLine Text:="        <vtkConfigurationsVersion>2.0</vtkConfigurationsVersion>"
    xmlFile.WriteLine Text:="        <projectName>" & projectName & "</projectName>"
    xmlFile.WriteLine Text:="    </info>"
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create Reference elements
    Dim ref As vtkReference
    For Each ref In cm.references
        ' create the configuration element
        xmlFile.WriteLine Text:="    <reference refID=""" & ref.id & """>"
        xmlFile.WriteLine Text:="        <name>" & ref.name & "</name>"
        If ref.GUID Like "" Then
            xmlFile.WriteLine Text:="        <path>" & ref.relPath & "</path>"
           Else
            xmlFile.WriteLine Text:="        <guid>" & ref.GUID & "</guid>"
        End If
        xmlFile.WriteLine Text:="    </reference>"
    Next
    
    ' Create Configuration elements
    Dim cf As vtkConfiguration, refList As String
    For Each cf In cm.configurations
        ' create the configuration element
        refList = ""
        For Each ref In cf.references
            If refList <> "" Then refList = refList & " "
            refList = refList & ref.id
        Next
        If refList = "" Then
            xmlFile.WriteLine Text:="    <configuration cID=""" & cf.id & """>"
           Else
            xmlFile.WriteLine Text:="    <configuration cID=""" & cf.id & """ refIDs=""" & refList & """>"
        End If
        xmlFile.WriteLine Text:="        <name>" & cf.name & "</name>"
        xmlFile.WriteLine Text:="        <path>" & escapedString(cf.genericPath) & "</path>"
        If cf.template <> "" Then _
        xmlFile.WriteLine Text:="        <templatePath>" & cf.template & "</templatePath>"
        xmlFile.WriteLine Text:="        <title>" & cf.projectName & "</title>" ' must be initialized in Workbook with Wb.BuiltinDocumentProperties("Title").Value
        xmlFile.WriteLine Text:="        <comment>" & cf.comment & "</comment>" ' must be initialized in Workbook with Wb.BuiltinDocumentProperties("Comments").Value
        If cf.password <> "" Then _
        xmlFile.WriteLine Text:="        <password>" & cf.password & "</password>"
        xmlFile.WriteLine Text:="    </configuration>"
    Next
    
    ' Create Module elements
    Dim md As vtkModule, modulePath As String
    For Each md In cm.modules
        xmlFile.WriteLine Text:="    <module mID=""" & md.id & """>"
        xmlFile.WriteLine Text:="        <name>" & md.name & "</name>"
        For Each cf In cm.configurations
            modulePath = md.getPathForConfiguration(confName:=cf.name)
            If Not modulePath Like "" Then xmlFile.WriteLine Text:="        <modulePath confId=""" & cf.id & """>" & modulePath & "</modulePath>"
        Next
        xmlFile.WriteLine Text:="    </module>"
    Next
    
    ' Close the file
    xmlFile.WriteLine Text:="</vtkConf>"
    xmlFile.Close
    
   On Error GoTo 0
   Exit Sub

vtkExportConfigurationsAsXML_Error:
    Dim s As String
    s = "vtkXMLutilities::exportConfigurationsAsXML"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "The " & projectName & "_DEV workbook is not opened"
        Case 76
            Err.Number = VTK_WRONG_FILE_PATH
            Err.Description = "The " & filePath & " path is unreachable"
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
            s = s & " -> " & Err.Source
    End Select

    Err.Raise Err.Number, s, Err.Description
End Sub

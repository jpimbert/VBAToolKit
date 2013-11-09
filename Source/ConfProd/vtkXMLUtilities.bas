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
    xmlFile.WriteLine Text:="<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""no""?>"
    xmlFile.WriteLine Text:="<!DOCTYPE vtkConf SYSTEM ""vtkConfigurationsDTD.dtd"">"
    xmlFile.WriteLine Text:="<vtkConf>"
    xmlFile.WriteBlankLines Lines:=1
    xmlFile.WriteLine Text:="    <info>"
    xmlFile.WriteLine Text:="        <vtkConfigurationsVersion>1.0</vtkConfigurationsVersion>"
    xmlFile.WriteLine Text:="        <projectName>" & projectName & "</projectName>"
    xmlFile.WriteLine Text:="    </info>"
    xmlFile.WriteBlankLines Lines:=1
    
    ' Create Configuration elements
    Dim cf As vtkConfiguration
    For Each cf In cm.configurations
        ' create the configuration element
        xmlFile.WriteLine Text:="    <configuration cID=""" & cf.ID & """>"
        xmlFile.WriteLine Text:="        <name>" & cf.name & "</name>"
        xmlFile.WriteLine Text:="        <path>" & cf.path & "</path>"
        xmlFile.WriteLine Text:="    </configuration>"
    Next
    
    ' Create Module elements
    Dim md As vtkModule, modulePath As String
    For Each md In cm.modules
        xmlFile.WriteLine Text:="    <module mID=""" & md.ID & """>"
        xmlFile.WriteLine Text:="        <name>" & md.name & "</name>"
        For Each cf In cm.configurations
            modulePath = md.getPathForConfiguration(confName:=cf.name)
            If Not modulePath Like "" Then xmlFile.WriteLine Text:="        <modulePath confId=""" & cf.ID & """>" & modulePath & "</modulePath>"
        Next
        xmlFile.WriteLine Text:="    </module>"
    Next
    
    ' Create References elements
    '   Designed only to be used with an active Workbook being a _DEV VBA project
    '   All references are exported to the XML configuration file
    '   except the VBAToolKit reference that is exported only for the _DEV configuration
    Dim r As VBIDE.Reference, allConfIDs As String, confIDsOnlyDEV As String
    ' Build ConfIDs lists
    For Each cf In cm.configurations
        If Not allConfIDs Like "" Then
            allConfIDs = allConfIDs & " " & cf.ID
           Else
            allConfIDs = """" & cf.ID
        End If
        If cf.isDEV Then
            If Not confIDsOnlyDEV Like "" Then
                confIDsOnlyDEV = confIDsOnlyDEV & " " & cf.ID
               Else
                confIDsOnlyDEV = """" & cf.ID
            End If
        End If
    Next
    allConfIDs = allConfIDs & """"
    confIDsOnlyDEV = confIDsOnlyDEV & """"
    ' Add reference elements to the XML configuration file
    For Each r In ActiveWorkbook.VBProject.References
        If r.name Like "VBAToolKit" Then
            xmlFile.WriteLine Text:="    <reference confIDs=" & confIDsOnlyDEV & ">"
           Else
            xmlFile.WriteLine Text:="    <reference confIDs=" & allConfIDs & ">"
        End If
        xmlFile.WriteLine Text:="        <name>" & r.name & "</name>"
        If r.GUID Like "" Then
            xmlFile.WriteLine Text:="        <path>" & r.fullPath & "</path>"
           Else
            xmlFile.WriteLine Text:="        <guid>" & r.GUID & "</guid>"
        End If
        xmlFile.WriteLine Text:="    </reference>"
    Next
    
   ' Create Title and Comments for the configuration
    
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

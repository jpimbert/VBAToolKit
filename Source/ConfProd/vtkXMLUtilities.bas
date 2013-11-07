Attribute VB_Name = "vtkXMLUtilities"
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
' Procedure : vtkExportAsXMLDOM
' Author    : Lucas Vitorino
' Purpose   : Export a DEV workbook as an XML DOM.
'
' Returns   : MSXML2.DOMDocument
' Raises    : - VTK_WORKBOOK_NOT_OPEN
'             - VTK_WORKBOOK_NOT_INITIALIZED
'             - VTK_PROJECT_NOT_INITIALIZED
'             - VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Function vtkExportAsXMLDOM(projectName As String) As MSXML2.DOMDocument
    Dim dom As MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    Dim rootNode As MSXML2.IXMLDOMElement
    Dim tmpEl As MSXML2.IXMLDOMElement
    Dim attr As MSXML2.IXMLDOMAttribute
    
    On Error GoTo vtkExportAsXMLDOM_Error
    
    ' If the project is not initialized
    Dim cm As vtkConfigurationManager
    Dim conf As vtkConfiguration
    Set cm = vtkConfigurationManagerForProject(projectName)
    If cm Is Nothing Then
        Err.Raise VTK_PROJECT_NOT_INITIALIZED
    End If
    
    Set dom = New MSXML2.DOMDocument
    Set node = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")
    dom.appendChild node

    ' The root node is an DOMElement, not a DOMNode
    Set rootNode = dom.createElement("vtkConf")
    
    With dom.appendChild(rootNode)
        
        ' The info element
        With .appendChild(dom.createElement("info"))
        
            ' Project name
            With .appendChild(dom.createElement("projectName"))
                .Text = projectName
            End With
            
            ' Version of vtkConfigurations
            With .appendChild(dom.createElement("vtkConfigurationsVersion"))
                .Text = "1.0"
            End With
            
        End With
        
        
        'The configuration elements
        For Each conf In cm.configurations
            With .appendChild(dom.createElement("configuration"))
                
                ' The name
                With .appendChild(dom.createElement("name"))
                    .Text = conf.name
                End With
                
                'The path
                With .appendChild(dom.createElement("path"))
                    .Text = conf.path
                End With
                
            End With
        Next
        
        
        ' The module elements
        Dim mo As vtkModule
        For Each mo In cm.modules
            With .appendChild(dom.createElement("module"))
                
                ' The name
                With .appendChild(dom.createElement("name"))
                    .Text = mo.name
                End With
                
                ' The path for each configuration
                For Each conf In cm.configurations
                    Set attr = dom.createAttribute("confName")
                    attr.NodeValue = conf.name
                    Set tmpEl = dom.createElement("path")
                    tmpEl.setAttributeNode attr
                    With .appendChild(tmpEl)
                        .Text = mo.getPathForConfiguration(conf.name)
                    End With
                Next
            
            End With
        Next
        
    End With

    On Error GoTo 0
    Set vtkExportAsXMLDOM = dom
    Exit Function

vtkExportAsXMLDOM_Error:
    Err.Source = "function vtkExportAsDOMXML of module vtkXMLutilities"
    
    Select Case Err.Number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_WORKBOOK_NOT_INITIALIZED
            Err.Description = "Workbook not initialized."
        Case -2147221080 ' Automation error undocumented by Microsoft
            Err.Number = VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_PROJECT_NOT_INITIALIZED
            Err.Description = "Project " & projectName & " has not been initialized. The name might be wrong."
        Case Else
            Debug.Print "Unexpected error " & Err.Number & " (" & Err.Description & ") in " & Err.Source
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number

    Exit Function
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkWriteXMLDOMToFile
' Author    : Lucas Vitorino
' Purpose   : - Write an XML DOM to a xml text file.
'             - The content of the file is nicely indented, to be human-readable.
'             - Overwrite the output file if it exists.
' Raises    : - VTK_DOM_NOT_INTIALIZED
'             - VTK_UNEXPECTED_ERRROR
' Notes     : Heavily based on code from Baptiste Wicht, http://baptiste-wicht.developpez.com/
'---------------------------------------------------------------------------------------
'
Public Sub vtkWriteXMLDOMToFile(dom As MSXML2.DOMDocument, filePath As String)

    Dim rdr As MSXML2.SAXXMLReader
    Dim wrt As MSXML2.MXXMLWriter
    
    On Error GoTo vtkWriteXMLDOMToFile_Error
    
    ' Check DOM intialization
    If dom Is Nothing Then
        Err.Raise VTK_DOM_NOT_INITIALIZED
    End If
    
    Set rdr = CreateObject("MSXML2.SAXXMLReader")
    Set wrt = CreateObject("MSXML2.MXXMLWriter")
    
    Dim oStream As ADODB.STREAM
    Set oStream = CreateObject("ADODB.STREAM")
    oStream.Open
    oStream.Charset = "ISO-8859-1"

    wrt.indent = True
    wrt.Encoding = "ISO-8859-1"
    wrt.output = oStream
    Set rdr.contentHandler = wrt
    Set rdr.errorHandler = wrt

    rdr.Parse dom
    wrt.flush

    oStream.SaveToFile filePath, adSaveCreateOverWrite

    On Error GoTo 0
    Exit Sub

vtkWriteXMLDOMToFile_Error:
    Err.Source = "function vtkWriteXMLDOMToFile of module vtkXMLutilities"
    
    Select Case Err.Number
        Case VTK_DOM_NOT_INITIALIZED
            Err.Description = "Dom object is not initialized."
        Case 3004 ' ADODB.Stream.SaveToFile failed because it couldn't find the path
            Err.Number = VTK_WRONG_FILE_PATH
            Err.Description = "File path is wrong. Make sure the folder tree is valid."
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number
    
    Exit Sub

End Sub

'---------------------------------------------------------------------------------------
' Procedure : exportConfigurationsAsXML
' Author    : Jean-Pierre IMBERT
' Date      : 07/11/2013
' Purpose   : Export the configurations of a project as a new XML file
'             - the "new" XML file format is the one designed for configuration management as XML
'             - this subroutine is temporary, dedicated to prepare the migration from
'               Excel sheet management of configurations to XML file management
' Raises    : - VTK_WORKBOOK_NOT_OPEN if the _DEV workbook containing the configuration sheet is not opened
'---------------------------------------------------------------------------------------
'
Public Sub exportConfigurationsAsXML(projectName As String, filePath As String)

   On Error GoTo exportConfigurationsAsXML_Error

    ' If the project is not initialized
    Dim cm As vtkConfigurationManager
    Dim conf As vtkConfiguration
    Set cm = vtkConfigurationManagerForProject(projectName)
    If cm Is Nothing Then
        Err.Raise VTK_WORKBOOK_NOT_OPEN
    End If
    

   On Error GoTo 0
   Exit Sub

exportConfigurationsAsXML_Error:
    
    Select Case Err.Number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "The " & projectName & "_DEV workbook is not opened"
        Case Else
            Err.Number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.Number, "vtkXMLutilities::exportConfigurationsAsXML -> " & Err.Source, Err.Description
End Sub

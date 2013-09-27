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
' Raises    : - VTK_WORBOOK_NOTOPEN
'             - VTK_UNEXPECTED_ERROR
'---------------------------------------------------------------------------------------
'
Public Function vtkExportAsXMLDOM(devWB As Workbook, projectName As String) As MSXML2.DOMDocument
    Dim dom As MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    Dim rootNode As MSXML2.IXMLDOMElement
    Dim tmpEl As MSXML2.IXMLDOMElement
    Dim attr As MSXML2.IXMLDOMAttribute
    
    On Error GoTo vtkExportAsXMLDOM_Error
'Debug.Print "1"
    ' If the workbook is not open
    If VtkWorkbookIsOpen(devWB.name) = False Then
        Err.Raise VTK_WORKBOOK_NOT_OPEN
    End If
'Debug.Print "2"
    Set dom = New MSXML2.DOMDocument
    Set node = dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""ISO-8859-1""")
    dom.appendChild node
'Debug.Print "3"
    ' The root node is an DOMElement, not a DOMNode
    Set rootNode = dom.createElement("vtkConf")
'Debug.Print "4"
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
        
        
        ' The reference elements
        Dim ref As reference
        For Each ref In devWB.VBProject.References
            With .appendChild(dom.createElement("reference"))
                
                ' The name
                With .appendChild(dom.createElement("name"))
                    .Text = ref.name
                End With
                
                'The GUID or the path if there is no GUID
                If ref.GUID = "" Then
                    With .appendChild(dom.createElement("path"))
                        .Text = ref.fullPath
                    End With
                Else
                    With .appendChild(dom.createElement("GUID"))
                        .Text = ref.GUID
                    End With
                End If
            
            End With
        Next
        
        
        'The configuration elements
        Dim cm As vtkConfigurationManager
        Set cm = vtkConfigurationManagerForProject(projectName)
        Dim conf As vtkConfiguration
        
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
'Debug.Print "6"
    Set vtkExportAsXMLDOM = dom
'Debug.Print "7"

    On Error GoTo 0
    Exit Function

vtkExportAsXMLDOM_Error:
    Err.source = "function vtkExportAsDOMXML of module vtkXMLutilities"
    
    Select Case Err.number
        Case VTK_WORKBOOK_NOTOPEN
            Err.number = VTK_WORBOOK_NOTOPEN
            Err.Description = "Workbook should be open."
        Case Else
            Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
            Err.number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.number
    
    Exit Function
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : vtkWriteXMLDOMToFile
' Author    : Lucas Vitorino
' Purpose   : - Write an XML DOM to a xml text file.
'             - The content of the file is nicely indented, to be human-readable.
'             - Overwrite the output file if it exists.
' Raises    : - VTK_UNEXPECTED_ERRROR
' Notes     : Heavily based on code from Baptiste Wicht, http://baptiste-wicht.developpez.com/
'---------------------------------------------------------------------------------------
'
Public Sub vtkWriteXMLDOMToFile(dom As MSXML2.DOMDocument, filePath As String)

    Dim rdr As MSXML2.SAXXMLReader
    Dim wrt As MSXML2.MXXMLWriter
    
    On Error GoTo vtkWriteXMLDOMToFile_Error

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
    Err.source = "function vtkWriteXMLDOMToFile of module vtkXMLutilities"
    Err.number = VTK_UNEXEPECTED_ERROR
    
    Err.Raise Err.number
    
    Exit Sub

End Sub

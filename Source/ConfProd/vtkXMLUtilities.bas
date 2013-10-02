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
    Err.source = "function vtkExportAsDOMXML of module vtkXMLutilities"
    
    Select Case Err.number
        Case VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_WORKBOOK_NOT_INITIALIZED
            Err.Description = "Workbook not initialized."
        Case -2147221080 ' Automation error undocumented by Microsoft
            Err.number = VTK_WORKBOOK_NOT_OPEN
            Err.Description = "Workbook should be open."
        Case VTK_PROJECT_NOT_INITIALIZED
            Err.Description = "Project " & projectName & " has not been initialized. The name might be wrong."
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
    Err.source = "function vtkWriteXMLDOMToFile of module vtkXMLutilities"
    
    Select Case Err.number
        Case VTK_DOM_NOT_INITIALIZED
            Err.Description = "Dom object is not initialized."
        Case 3004 ' ADODB.Stream.SaveToFile failed because it couldn't find the path
            Err.number = VTK_WRONG_FILE_PATH
            Err.Description = "File path is wrong. Make sure the folder tree is valid."
        Case Else
            Err.number = VTK_UNEXPECTED_ERROR
    End Select

    Err.Raise Err.number
    
    Exit Sub

End Sub



'---------------------------------------------------------------------------------------
' Procedure : vtkCountElementsInNode
' Author    : Lucas Vitorino
' Purpose   : Counts the number of elements (sub nodes) named elementName in a XML DOM Node.
' Notes     : The search is only one-level deep.
' Raises    : - VTK_UNEXPECTED_ERROR
'             - VTK_OJBECT_NOT_INITIALIZED
'---------------------------------------------------------------------------------------
'
Public Function vtkCountElementsInNode(node As MSXML2.IXMLDOMNode, elementName As String) As Integer
    
    Dim count As Integer: count = 0
    
    On Error GoTo vtkCountElementsInNode_Error

    Dim subNode As MSXML2.IXMLDOMNode
    For Each subNode In node.ChildNodes
        If StrComp(subNode.BaseName, elementName) = 0 Then count = count + 1
    Next
    
    vtkCountElementsInNode = count

    On Error GoTo 0
    Exit Function

vtkCountElementsInNode_Error:
    Err.source = "vtkCountElementsInNode of module vtkXMLUtilities"
    
    Select Case Err.number
        Case VTK_OBJECT_NOT_INITIALIZED
            Err.Description = "The object passed as a parameter is set to nothing."
        Case Else
            Err.number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.number, Err.source, Err.Description

    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkCountElementsInDom
' Author    : Lucas Vitorino
' Purpose   : Counts the number of elements (sub nodes) named elementName in the root node of an XML DOM.
' Notes     : The search is only one-level deep.
' Raises    : - VTK_UNEXPECTED_ERROR
'             - VTK_OJBECT_NOT_INITIALIZED
'---------------------------------------------------------------------------------------
'
Public Function vtkCountElementsInDom(dom As MSXML2.DOMDocument, elementName As String) As Integer

    On Error GoTo vtkCountElementsInDom_Error
    
    If dom Is Nothing Then
        Err.Raise VTK_OBJECT_NOT_INITIALIZED
    End If
    
    Dim rootNode As MSXML2.IXMLDOMNode
    Set rootNode = dom.ChildNodes.Item(1)
    
    vtkCountElementsInDom = vtkCountElementsInNode(rootNode, elementName)
    
    On Error GoTo 0
    Exit Function

vtkCountElementsInDom_Error:
    Err.source = "vtkCountElementsInDom of module vtkXMLUtilities"
    
    Select Case Err.number
        Case VTK_OBJECT_NOT_INITIALIZED
            Err.Description = "The object passed as a parameter is set to Nothing."
        Case Else
            Err.number = VTK_UNEXPECTED_ERROR
    End Select
    
    Err.Raise Err.number, Err.source, Err.Description

    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : vtkGetElementsTextInNode
' Author    : Lucas Vitorino
' Purpose   : Returns the text of the elements that have the given name
'             (and optionally the given attribute) immediatly below the given node.
' Notes     : The search is only one-level deep.
' Raises    : - VTK_UNEXPECTED_ERROR
'             - VTK_OBJECT_NOT_INITIALIZED
'---------------------------------------------------------------------------------------
'
Public Function vtkGetElementTextInNode(node As MSXML2.IXMLDOMNode, _
                                        elementName As String, _
                                        index As Integer _
                                        ) As String
                                              
    On Error GoTo vtkGetElementTextInNode_Error
    
    Dim count As Integer: count = 0
    Dim subNode As MSXML2.IXMLDOMNode

    ' Go through all the subnodes
    For Each subNode In node.ChildNodes
        ' If a matching subnode is found
        If subNode.BaseName = elementName Then
            ' If it corresponds to the desired index, return it
            count = count + 1
            If count = index Then
                vtkGetElementTextInNode = subNode.Text
                Exit Function
            End If
        End If
    Next

    ' If nothing has been found, we raise an error,
    ' as the empty string "" may be a return value.
    Err.Raise VTK_ELEMENT_NOT_FOUND

    On Error GoTo 0
    Exit Function

vtkGetElementTextInNode_Error:
    Err.source = "vtkGetElementTextInNode of module vtkXMLUtilities"
    
    Select Case Err.number
        Case VTK_OBJECT_NOT_INITIALIZED
            Err.Description = "The object passed as a parameter is set to Nothing."
        Case VTK_ELEMENT_NOT_FOUND
            Err.Description = "The element has not been found. Either the name is wrong, or the index is out of bounds."
        Case Else
            Err.number = VTK_UNEXPECTED_ERROR
    End Select
    
    Debug.Print "Unexpected error " & Err.number & " (" & Err.Description & ") in " & Err.source
    Err.Raise Err.number, Err.source, Err.Description

    Exit Function
End Function


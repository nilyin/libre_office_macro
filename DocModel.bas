REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option VBASupport 1

' Constants for document processing
Const STYLE_HEAD = "Heading"  ' Prefix for heading paragraph styles
Const CODE_LINE_NUM = True    ' Enable line numbering in code blocks

' Enumeration for different types of document nodes
Enum NodeType
    Section = 1    ' Document sections
    Style = 2      ' Paragraph styles
    List = 3       ' List structures
    Table = 4      ' Table structures
    Paragraph = 5  ' Individual paragraphs
End Enum

' Structure representing a node in the document tree
Type Node
    type_ As NodeType     ' Type of node (Section, Style, List, Table, Paragraph)
    value As Variant      ' LibreOffice object for Paragraph and Table nodes
    name_ As String       ' Name identifier for Section and Style nodes
    children As Variant   ' Collection of child nodes for Section and Style nodes
    level As Integer      ' Nesting level in document hierarchy
End Type

' Enumeration for section processing states
Enum SectionState
    End_ = 1       ' End current section
    New_ = 2       ' Start new section
    Continue = 3   ' Continue current section
End Enum

Function SectionTest(ByRef curNode, ByRef curPara, ByRef sectionNames) As SectionState
    If IsEmpty(curPara.TextSection) Then
        If curNode.level > 0 Then
            SectionTest = SectionState.End_
            Exit Function
        EndIf
    ElseIf IsEmpty(curNode.name_) Or curPara.TextSection.Name <> curNode.name_ Then
        Dim sectionName : sectionName = Null
        If Not IsEmpty(curNode.name_) Then
            On Error Resume Next
            sectionName = sectionNames.Item(curPara.TextSection.Name)
            If Not IsNull(sectionName) Then
                SectionTest = SectionState.End_
                Exit Function
            End If
        End If
        SectionTest = SectionState.New_
        Exit Function
    End If
    SectionTest = SectionState.Continue
End Function

Function MakeNewSection(ByRef curNode, ByRef curPara, ByRef sectionNames) As Node
    sectionNames.Add(True, curPara.TextSection.Name)
    Dim newSec As Node
    With newSec
        .type_ = NodeType.Section
        .name_ = curPara.TextSection.Name
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(newSec)
    MakeNewSection = newSec
End Function

Function GetNodeStyle(ByRef curNode, ByRef curPara) As Node
    Dim i As Integer : i = curNode.children.Count
    If i > 0 Then
        Dim lastItem : lastItem = curNode.children.Item(i)
        If lastItem.type_ = NodeType.Style And _
            lastItem.name_ = curPara.ParaStyleName Then
            GetNodeStyle = lastItem
            Exit Function
        End If
    End If
    Dim nodeStyle As Node
    With nodeStyle
        .type_ = NodeType.Style
        .name_ = curPara.ParaStyleName
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(nodeStyle)
    GetNodeStyle = nodeStyle
End Function

Function GetNodeList(ByRef curNode, ByRef curPara) As Node
    Dim i As Integer : i = curNode.children.Count
    If i > 0 Then
        Dim lastItem : lastItem = curNode.children.Item(i)
        If lastItem.type_ = NodeType.List Then
            i = lastItem.children.Count
            If i > 0 Then
                Dim listItem
                listItem = lastItem.children.Item(i)
                If listItem.type_  = NodeType.Paragraph Then
                    If listItem.value.NumberingLevel = curPara.NumberingLevel Then
                        GetNodeList = lastItem
                        Exit Function
                    End If
                Else
                    GetNodeList = GetNodeList(lastItem, curPara)
                    Exit Function
                End If
            End If
            GetNodeList = GetNodeList(lastItem, curPara)
            Exit Function
        End If
    End If
    Dim nodeList As Node
    With nodeList
        .type_ = NodeType.List
        .name_ = IIf(curPara.ListLabelString = "", "Marked", "Numbered")
        .level = curNode.level + 1
        .children = New Collection
    End With
    curNode.children.Add(nodeList)
    GetNodeList = nodeList
End Function

Sub SectionParse(ByRef paraEnum, ByRef curPara, ByRef curNode, ByRef sectionNames)
    ' Enumerate paragraphs, include tables
    Do
' emulate key word "continue" in C++
continue:
		    
        ' Process the tables
        If curPara.supportsService("com.sun.star.text.TextTable") Then
    	    Dim nodeTable As Node
            With nodeTable
                .type_ = NodeType.Table
                .level = curNode.level + 1
                .value = curPara
            End With
            curNode.children.Add(nodeTable)  
    	    
        ' Process the paragrath
        Elseif curPara.supportsService("com.sun.star.text.Paragraph") Then
            Dim secState As SectionState
            secState = SectionTest(curNode, curPara, sectionNames)            
            Select Case secState 
                Case SectionState.End_
                    Exit Sub
                Case SectionState.New_
                    Dim newSec As Node
                    newSec = MakeNewSection(curNode, curPara, sectionNames)
                    SectionParse paraEnum, curPara, newSec, sectionNames
                    GoTo continue
            End Select
            
            ' Process the Style
            Dim nodeStyle As Node
            nodeStyle = GetNodeStyle(curNode, curPara)           
            Dim nodePara As Node
            With nodePara
                .type_ = NodeType.Paragraph
                .level = nodeStyle.level + 1
                .value = curPara
            End With
            If curPara.NumberingIsNumber And _
                Left(curPara.ParaStyleName, 7) <> STYLE_HEAD Then
                Dim nodeList As Node
                nodeList = GetNodeList(nodeStyle, curPara)
                nodePara.level = nodeList.level + 1
                nodeList.children.Add(nodePara)
            Else
                nodeStyle.children.Add(nodePara)
            End If
      
        End If
        If Not paraEnum.hasMoreElements() Then Exit Do
        curPara = paraEnum.nextElement()
    Loop
End Sub

Sub ExportToFile (ByRef text_ As String, Comp As Object, Optional suffix As Variant)
    If IsMissing(suffix) Then suffix = "_export.txt"
    Dim FileNo As Integer, Filename As String
    Filename = convertToURL(replace(convertFromURL(Comp.URL), ".odt", suffix))
	FileNo = Freefile
	Open Filename For Output As #FileNo
	Print #FileNo, text_
End Sub

Function MakeModel(ByRef Comp As Object) As Node
    Dim sectionNames As New Collection
    Dim docTree As Node
    With docTree
        .type_ = NodeType.Section
        .level = 0
        .children = New Collection
    End With
 
    ' Enumerate paragraphs, include tables
    Dim paraEnum : paraEnum = Comp.getText().createEnumeration()
    Dim curPara
    If paraEnum.hasMoreElements() Then
        curPara = paraEnum.nextElement()
        SectionParse paraEnum, curPara, docTree, sectionNames
    End If
    MakeModel = docTree
End Function

Sub MakeDocHtmlView(Optional Comp As Variant)
    Dim doc As Object : doc = ThisComponent
    If Not IsMissing(Comp) Then
        If Not IsEmpty(Comp) Then doc = Comp
    End If

    Dim dView As New DocView : dView = New DocView
    Dim vHtml As New ViewHtml : vHtml = New ViewHtml
    vHtml.docView = dView
    dView.docTree = MakeModel(doc)
    dView.viewAdapter = vHtml
    dView.props = New Collection
    With dView.props
        .Add(CODE_LINE_NUM, "CodeLineNum") ' Enumerate code lines 1, 2, 3 ... n
    End With
    ExportToFile dView.MakeView(), doc, "_export.html"
End Sub

Sub MakeDocHfmView(Optional Comp As Variant)
    Dim doc As Object : doc = ThisComponent
    If Not IsMissing(Comp) Then
        If Not IsEmpty(Comp) Then doc = Comp
    End If

    Dim dView As New DocView : dView = New DocView
    Dim vHfm As New ViewHfm : vHfm = New ViewHfm
    vHfm.docView = dView
    dView.docTree = MakeModel(doc)
    dView.viewAdapter = vHfm
    dView.props = New Collection
    With dView.props
        .Add(CODE_LINE_NUM, "CodeLineNum") ' Enumerate code lines 1, 2, 3 ... n
    End With
    ExportToFile dView.MakeView(), doc, "_export_hfm.txt"
End Sub

' "C:\Program Files\LibreOffice\program\soffice.exe" --invisible --nofirststartwizard --headless --norestore macro:///DocExport.DocModel.ExportDir("D:\cpp\habr\002-hfm",0)
Sub ExportDir(Folder As String, Optional Hfm As Variant)
    Dim useHfm As Boolean : useHfm = True
    If Not IsMissing(Hfm) Then
        If IsNumeric(Hfm) Then useHfm = CBool(Hfm)
        If VarType(Hfm) = vbBoolean Then useHfm = Hfm
    End If  
    Dim Props(0) as New com.sun.star.beans.PropertyValue
    Props(0).NAME = "Hidden" 
    Props(0).Value = True 
    Dim Comp As Object
    Dim url, fname As String : fname = Dir$(Folder + "\" + "*.odt", 0)    
    Do
        url = ConvertToUrl(Folder + "\" + fname)
        Comp = StarDesktop.loadComponentFromURL(url, "_blank", 0, Props)
        If useHfm Then
            MakeDocHfmView Comp
        Else
            MakeDocHtmlView Comp
        End If
        fname = Dir$
        call Comp.close(True)
    Loop Until fname = ""
End Sub


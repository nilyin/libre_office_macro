REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

' Constants for document style recognition
Const STYLE_QUOT = "Quotations"  ' Style name for quotation blocks
Const STYLE_CODE = "code_"       ' Prefix for code block styles
Const STYLE_HEAD = "Heading"     ' Prefix for heading styles

' Public variables for document processing
Public docTree                   ' Document tree structure containing parsed content
Public viewAdapter              ' Output format adapter (HFM, HTML, etc.)
Public props                    ' Properties collection for formatting options

' Process document node based on its style type
' @param node: Document node with style information
' @return: Formatted content based on style type
Function PrintNodeStyle(ByRef node)
    Dim s As String : s = ""  ' Result string
    ' Handle different style types
    If node.name_ = STYLE_QUOT Then
        s = viewAdapter.Quote(node)  ' Format as blockquote
    ElseIf Left(node.name_, 5) = STYLE_CODE Then
        s = viewAdapter.Code(node)   ' Format as code block
    ElseIf Left(node.name_, 7) = STYLE_HEAD Then
        ' Process heading with potential bookmarks
        Dim textPortion, enumPortion
        enumPortion = node.children(1).value.createEnumeration()
        Do While enumPortion.hasMoreElements()
            textPortion = enumPortion.nextElement()
            If textPortion.TextPortionType = "Text" Then
                s = s & viewAdapter.Head(node)     ' Add heading text
            ElseIf textPortion.TextPortionType = "Bookmark" Then
                s = s & viewAdapter.Anchor(textPortion)  ' Add bookmark anchor
            End If
        Loop
    Else
        s = viewAdapter.ParaStyle(node)  ' Default paragraph style
    End If
    PrintNodeStyle = s
End Function

' Process LibreOffice paragraph object and extract all content
' @param oPara: LibreOffice paragraph object
' @param level: Nesting level for formatting
' @param lineNum: Optional line number for code blocks
' @return: Formatted paragraph content
Function PrintNodeParaLO(ByRef oPara, level As Long, Optional lineNum As Variant)
    ' LibreOffice service type constants
    Dim textGraphObj$ : textGraphObj$ = "com.sun.star.text.TextGraphicObject"
    Dim drawShape$ : drawShape$ = "com.sun.star.drawing.Shape"
    Dim textEmbObj$ : textEmbObj$ = "com.sun.star.text.TextEmbeddedObject"
    
    ' Process graphics anchored to paragraph (enumerate as TextContent)
    Dim contEnum : contEnum = _
        oPara.createContentEnumeration("com.sun.star.text.TextContent")
    Dim curContent, s As String : s = ""  ' Current content and result string
    
    ' Add line number prefix for code blocks - fix for newer LibreOffice versions
    If Not IsMissing(lineNum) Then
        If IsNumeric(lineNum) And lineNum > 0 Then
            s = s & Format_Num(CInt(lineNum)) & " "
        End If
    End If
    
    ' Process paragraph-anchored content
    Do While contEnum.hasMoreElements()
        curContent = contEnum.nextElement()           
        If curContent.supportsService(textGraphObj$) Then
            s = s & viewAdapter.Image(curContent)  ' Process paragraph-anchored images
        ElseIf curContent.supportsService(drawShape$) Then
            ' Drawing shapes anchored to paragraph (currently not processed)
        End If
    Loop
    
    ' Process character-anchored graphics and inline content
    ' These are enumerated as TextPortionType within the paragraph
    Dim textPortion, enumPortion : enumPortion = oPara.createEnumeration()
    ' Process each text portion in the paragraph
    Do While enumPortion.hasMoreElements()
        textPortion = enumPortion.nextElement()
        If textPortion.TextPortionType = "Text" Then
            ' Process text with formatting and hyperlinks
            If Not IsEmpty(textPortion.HyperLinkURL) And _
                textPortion.HyperLinkURL <> "" Then
			    s = s & viewAdapter.Link(textPortion)  ' Format as hyperlink
      	    ElseIf textPortion.CharWeight = com.sun.star.awt.FontWeight.BOLD Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Bold")  ' Bold formatting
      	    ElseIf textPortion.CharPosture = com.sun.star.awt.FontSlant.ITALIC Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Italic")  ' Italic formatting
      	    ElseIf textPortion.CharUnderline = com.sun.star.awt.FontUnderline.SINGLE Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Underline")  ' Underline formatting
      	    ElseIf textPortion.CharStrikeout = com.sun.star.awt.FontStrikeout.SINGLE Then
			    s = s & viewAdapter.FontDecorate(textPortion, "Strikeout")  ' Strikethrough formatting
            Else
			    s = s & textPortion.String  ' Plain text
			End If
        ElseIf textPortion.TextPortionType = "Frame" Then
            ' Process inline frames (images, shapes, formulas)
            Dim framePortion, enumFrame
            enumFrame = textPortion.createContentEnumeration(textGraphObj$)
            Do While enumFrame.hasMoreElements()
                framePortion = enumFrame.nextElement()
                If framePortion.supportsService(textGraphObj$) Then
                    s = s & viewAdapter.InlineImage(framePortion)  ' Inline images
                ElseIf framePortion.supportsService(drawShape$) Then
                    ' Inline drawing shapes (currently not processed)
                ElseIf framePortion.supportsService(textEmbObj$) And _
                    framePortion.FrameStyleName = "Formula" Then
                    s = s & viewAdapter.Formula(framePortion.Component.Formula)  ' Math formulas
                End If
            Loop
        ElseIf textPortion.TextPortionType = "Bookmark" Then
            s = s & viewAdapter.Anchor(textPortion)  ' Process bookmarks as anchors
        End If
    Loop
    ' Apply final formatting based on paragraph type
    If oPara.NumberingIsNumber Then
        PrintNodeParaLO = viewAdapter.FormatList(oPara, s, level)  ' Format as list item
    ElseIf Not IsMissing(lineNum) Then
        If IsNumeric(lineNum) Then
            If CInt(lineNum) = 0 Then
                PrintNodeParaLO = viewAdapter.FormatPara(s, level, 0)  ' Code block without extra line break
            ElseIf CInt(lineNum) > 0 Then
                PrintNodeParaLO = s & CHR$(10)  ' Code line with line break
            Else
                PrintNodeParaLO = viewAdapter.FormatPara(s, level, 1)  ' Regular paragraph with line break
            End If
        Else
            PrintNodeParaLO = viewAdapter.FormatPara(s, level, 1)  ' Regular paragraph with line break
        End If
    Else
        PrintNodeParaLO = viewAdapter.FormatPara(s, level, 1)  ' Regular paragraph with line break
    End If
End Function

' Process paragraph node wrapper
' @param nodePara: Document paragraph node
' @param lineNum: Optional line number for code blocks
' @return: Formatted paragraph content
Function PrintNodePara(ByRef nodePara, Optional lineNum As Variant)
    If IsMissing(lineNum) Then
        PrintNodePara = PrintNodeParaLO(nodePara.value, nodePara.level)
        Exit Function
    End If
    PrintNodePara = PrintNodeParaLO(nodePara.value, nodePara.level, lineNum)
End Function

' Process table node and format as markdown table
' @param nodeTable: Document table node
' @return: Formatted markdown table
Function PrintNodeTable(ByRef nodeTable)
    ' Variables for table processing
    Dim oTable, oCell, oText, oEnum, oPar, t, r, c
    Dim nRow As Long, nCol As Long, Rows As Long, Colls As Long
    
    oTable = nodeTable.value : t = ""  ' Get LibreOffice table object
    Rows = oTable.getRows().getCount() - 1     ' Get row count (0-based)
    Colls = oTable.getColumns().getCount() - 1 ' Get column count (0-based)
    
    ' Process each row
    For nRow = 0 To Rows
        r = ""  ' Row content string
        ' Process each cell in the row
        For nCol = 0 To Colls
            oCell = oTable.getCellByPosition(nCol, nRow)  ' Get cell object
            oText = oCell.getText()  ' Get cell text object
            c = ""  ' Cell content string
            
            ' Process all paragraphs in the cell
            oEnum = oText.createEnumeration()
            Do While oEnum.hasMoreElements()
                oPar = oEnum.nextElement()
                If oPar.supportsService("com.sun.star.text.Paragraph") Then
                    c = c & PrintNodeParaLO(oPar, nodeTable.level + 3, 0)  ' Process cell paragraph
                End If
            Loop
            r = r & viewAdapter.FormatCell(c, nodeTable.level + 2, nCol, nRow)  ' Format cell
        Next
        t = t & viewAdapter.FormatRow(r, nodeTable.level + 1, nRow, Colls)  ' Format row
    Next
    PrintNodeTable = viewAdapter.FormatTable(t, nodeTable.level)  ' Format complete table
End Function

Function PrintTree(ByRef node, Optional ByRef props As Collection)
    Dim child, lineNum : lineNum = 0
    Dim s : s = ""
    ' Fix for newer LibreOffice versions - safe Collection parameter checking
    On Error Resume Next
    If Not IsMissing(props) Then
        If TypeName(props) = "Collection" Then
            If props("CodeLineNum") Then lineNum = 1
        End If
    End If
    On Error GoTo 0
    For Each child In node.children
        If child.type_ = NodeType.Section Then
            s = s & viewAdapter.Section(child)
        ElseIf child.type_ = NodeType.Style Then
            s = s & PrintNodeStyle(child)
        ElseIf child.type_ = NodeType.List Then
            s = s & viewAdapter.List(child)
        ElseIf child.type_ = NodeType.Paragraph Then
            If lineNum > 0 Then
                s = s & PrintNodePara(child, lineNum)
                lineNum = lineNum + 1
            Else
                s = s & PrintNodePara(child)
            End If
        ElseIf child.type_ = NodeType.Table Then
            s = s & PrintNodeTable(child)
        End If
    Next
    PrintTree = s
End Function

' Process image with copying logic for embedded images
' @param imageObj: LibreOffice image object
' @param docURL: Document URL for determining target directory
' @return: Formatted markdown image string
Public Function ProcessImage(ByRef imageObj, ByRef docURL As String) As String
    Dim altText As String : altText = IIf(imageObj.Title = "", "image", imageObj.Title)
    Dim imageName As String
    On Error Resume Next
    imageName = imageObj.Graphic.OriginURL
    If imageName = "" Then imageName = imageObj.GraphicURL
    On Error GoTo 0
    
    If imageName <> "" Then
        ' Check if it's a remote URL
        If Left(LCase(imageName), 4) = "http" Then
            ProcessImage = "![" & altText & "](" & imageName & ")"
        Else
            ' Extract and copy embedded image
            Dim fileName As String : fileName = Mid(imageName, InStrRev(imageName, "/") + 1)
            fileName = LCase(fileName)
            fileName = Replace(fileName, "(", "-")
            fileName = Replace(fileName, ")", "")
            fileName = Replace(fileName, " ", "-")
            
            Dim docDir As String : docDir = Left(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), "\"))
            CopyImageFile imageName, docDir, fileName
            ProcessImage = "![" & altText & "](./img/" & fileName & ")"
        End If
    Else
        ProcessImage = "![" & altText & "](./img/missing-image.png)"
    End If
End Function

' Generate the complete formatted output from document tree
' @return: Complete formatted document content
Public Function MakeView() As String
    MakeView = PrintTree(docTree)
End Function


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
Public imageCounter As Long      ' Counter for image naming
Public docPrefix As String       ' Document name prefix for image naming
Public docURL As String          ' Document URL for image processing

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
        If node.children.Count > 0 Then
            Dim textPortion, enumPortion
            On Error Resume Next
            enumPortion = node.children(1).value.createEnumeration()
            If Err.Number = 0 Then
                Do While enumPortion.hasMoreElements()
                    textPortion = enumPortion.nextElement()
                    If textPortion.TextPortionType = "Text" Then
                        s = s & viewAdapter.Head(node)     ' Add heading text
                    ElseIf textPortion.TextPortionType = "Bookmark" Then
                        s = s & viewAdapter.Anchor(textPortion)  ' Add bookmark anchor
                    End If
                Loop
            Else
                s = s & viewAdapter.Head(node)  ' Fallback to simple heading
            End If
            On Error GoTo 0
        Else
            s = s & viewAdapter.Head(node)  ' Fallback to simple heading
        End If
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
                PrintNodeParaLO = s & "  " & CHR$(10)  ' Code line with line break
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

' Format number with leading zeros (minimum 2 digits)
' @param num: Number to format
' @return: Formatted number string
Private Function Format_Num(ByVal num As Long) As String
    If num < 10 Then
        Format_Num = "0" & CStr(num)
    Else
        Format_Num = CStr(num)
    End If
End Function

' Generate document prefix from filename for image naming
' @param docURL: Document URL
' @return: Prefix string for image names
Private Function GenerateDocPrefix(ByRef docURL As String) As String
    Dim fileName As String : fileName = Mid(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), GetPathSeparator()) + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1) ' Remove extension
    
    ' Split by separators and take first 4 chars from each word
    Dim words : words = Split(Replace(Replace(fileName, "_", " "), "-", " "), " ")
    Dim prefix As String : prefix = ""
    Dim i As Long
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            If Len(words(i)) >= 4 Then
                prefix = prefix & Left(words(i), 4) & "_"
            Else
                prefix = prefix & words(i) & "_"
            End If
        End If
    Next
    If Len(prefix) > 0 Then prefix = Left(prefix, Len(prefix) - 1) ' Remove trailing underscore
    GenerateDocPrefix = prefix
End Function

' Generate image folder name from document URL
' @param docURL: Document URL
' @return: Image folder name with pattern "img_" + source filename
Private Function GenerateImageFolderName(ByRef docURL As String) As String
    Dim fileName As String : fileName = Mid(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), GetPathSeparator()) + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1) ' Remove extension
    GenerateImageFolderName = "img_" & fileName
End Function

' Extract and save image from LibreOffice graphic object
' @param imageObj: LibreOffice image object
' @param targetDir: Target directory path
' @param fileName: Target filename
' @return: True if extraction successful
Private Function ExtractImageFile(ByRef imageObj, ByRef targetDir As String, ByRef fileName As String, ByRef docURL As String) As Boolean
    On Error Resume Next
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    Dim imgFolderName As String : imgFolderName = GenerateImageFolderName(docURL)
    Dim imgDir As String : imgDir = targetDir & imgFolderName
    Dim targetPath As String : targetPath = imgDir & GetPathSeparator() & fileName
    
    ' Create img directory if it doesn't exist
    If Not fso.FolderExists(imgDir) Then fso.CreateFolder(imgDir)
    
    ' Try to get the graphic object and export it
    Dim graphic : graphic = imageObj.Graphic
    If Not IsEmpty(graphic) Then
        ' Create GraphicProvider service
        Dim graphicProvider : graphicProvider = CreateUnoService("com.sun.star.graphic.GraphicProvider")
        
        ' Set up export properties
        Dim exportProps(1) As New com.sun.star.beans.PropertyValue
        exportProps(0).Name = "URL"
        exportProps(0).Value = ConvertToURL(targetPath)
        exportProps(1).Name = "MimeType"
        exportProps(1).Value = "image/png"
        
        ' Export the graphic
        graphicProvider.storeGraphic(graphic, exportProps())
        ExtractImageFile = (Err.Number = 0)
    Else
        ExtractImageFile = False
    End If
    On Error GoTo 0
End Function

' Copy image file from external source to img folder
' @param sourceURL: Source image URL
' @param targetDir: Target directory path
' @param fileName: Target filename
' @return: True if copy successful
Private Function CopyImageFile(ByRef sourceURL As String, ByRef targetDir As String, ByRef fileName As String, ByRef docURL As String) As Boolean
    On Error Resume Next
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    Dim imgFolderName As String : imgFolderName = GenerateImageFolderName(docURL)
    Dim imgDir As String : imgDir = targetDir & imgFolderName
    Dim targetPath As String : targetPath = imgDir & GetPathSeparator() & fileName
    
    ' Create img directory if it doesn't exist
    If Not fso.FolderExists(imgDir) Then fso.CreateFolder(imgDir)
    
    ' External file reference
    Dim sourcePath As String : sourcePath = ConvertFromURL(sourceURL)
    If fso.FileExists(sourcePath) Then
        fso.CopyFile sourcePath, targetPath, True
        CopyImageFile = (Err.Number = 0)
    Else
        CopyImageFile = False
    End If
    On Error GoTo 0
End Function

' Process image with copying logic for embedded images
' @param imageObj: LibreOffice image object
' @param docURL: Document URL for determining target directory
' @return: Formatted markdown image string
Public Function ProcessImage(ByRef imageObj, ByRef docURL As String) As String
    ' Initialize document prefix if not set
    If docPrefix = "" Then docPrefix = GenerateDocPrefix(docURL)
    
    Dim altText As String : altText = IIf(imageObj.Title = "", "image", imageObj.Title)
    Dim imageName As String
    On Error Resume Next
    imageName = imageObj.Graphic.OriginURL
    If imageName = "" Then imageName = imageObj.GraphicURL
    On Error GoTo 0
    
    ' Check if image has a hyperlink URL (external link)
    Dim hasExternalLink As Boolean : hasExternalLink = False
    On Error Resume Next
    If imageObj.HyperLinkURL <> "" Then
        If Left(LCase(imageObj.HyperLinkURL), 4) = "http" Then
            hasExternalLink = True
            ProcessImage = "![" & altText & "](" & imageObj.HyperLinkURL & ")"
        End If
    End If
    On Error GoTo 0
    
    If hasExternalLink Then
        Exit Function
    End If
    
    ' Check if it's a remote URL in the image source itself
    If imageName <> "" And Left(LCase(imageName), 4) = "http" Then
        ProcessImage = "![" & altText & "](" & imageName & ")"
        Exit Function
    End If
    
    ' This is an embedded/local image - extract it
    imageCounter = imageCounter + 1
    
    ' Generate filename based on requirements
    Dim fileName As String
    If imageObj.Name <> "" And imageObj.Name <> "Graphic1" And imageObj.Name <> "Image1" Then
        ' Use existing name if available and not default
        fileName = imageObj.Name
        ' Ensure proper extension
        If Right(LCase(fileName), 4) <> ".png" And Right(LCase(fileName), 4) <> ".jpg" And Right(LCase(fileName), 5) <> ".jpeg" Then
            fileName = fileName & ".png"
        End If
    Else
        ' Generate name: prefix_XX.png
        fileName = docPrefix & "_" & Format_Num(imageCounter) & ".png"
    End If
    
    Dim docDir As String : docDir = Left(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), GetPathSeparator()))
    
    ' Try to extract embedded image first, then try copying external file
    Dim success As Boolean : success = False
    If Not IsEmpty(imageObj.Graphic) Then
        success = ExtractImageFile(imageObj, docDir, fileName, docURL)
    End If
    
    If Not success And imageName <> "" Then
        success = CopyImageFile(imageName, docDir, fileName, docURL)
    End If
    
    ' Generate dynamic folder name for markdown reference
    Dim imgFolderName As String : imgFolderName = GenerateImageFolderName(docURL)
    ' Always return the markdown reference, even if extraction failed
    ProcessImage = "![" & altText & "](" & imgFolderName & "/" & fileName & ")"
End Function

' Generate the complete formatted output from document tree
' @return: Complete formatted document content
Public Function MakeView() As String
    MakeView = PrintTree(docTree)
End Function


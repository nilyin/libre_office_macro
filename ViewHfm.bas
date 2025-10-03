REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit
Option Compatible
Option ClassModule

' Constants for section and style handling
Const SEC_HEADING = ""              ' Empty string identifier for main heading sections
Const STYLE_HEADING = "Contents"     ' Style name for content headings
Const SHIFT_CNT = 4                  ' Number of spaces for list indentation

' Public variables
Public docView                       ' Reference to the main document view object
Public fontStyles As New Collection  ' Collection storing font decoration patterns

' Initialize font decoration patterns for markdown formatting
Private Sub Class_Initialize()
    fontStyles.Add(Array("**", "**"), "Bold")        ' Bold text markers
    fontStyles.Add(Array("*", "*"), "Italic")        ' Italic text markers
    fontStyles.Add(Array("<u>", "</u>"), "Underline") ' Underline HTML tags
    fontStyles.Add(Array("~~", "~~"), "Strikeout")   ' Strikethrough markers
End Sub

' Generate TOC anchor based on hybrid approach
' @param headingText: The heading text to convert to anchor
' @return: Anchor string for TOC links
Private Function GenerateTOCAnchor(ByRef headingText As String) As String
    Dim cleanText As String : cleanText = LCase(headingText)
    
    ' Extract numbering part (everything before first space)
    Dim spacePos As Integer : spacePos = InStr(cleanText, " ")
    Dim numbering As String : numbering = ""
    Dim titleText As String : titleText = cleanText
    
    If spacePos > 0 Then
        numbering = Left(cleanText, spacePos - 1)
        titleText = Mid(cleanText, spacePos + 1)
    End If
    
    ' Clean title text: remove dots and clean up
    Dim cleanTitle As String : cleanTitle = Replace(titleText, ".", "")
    cleanTitle = Replace(cleanTitle, " ", "-")
    cleanTitle = Replace(cleanTitle, ",", "")
    cleanTitle = Replace(cleanTitle, "(", "")
    cleanTitle = Replace(cleanTitle, ")", "")
    cleanTitle = Replace(cleanTitle, ":", "")
    ' Remove consecutive dashes
    Do While InStr(cleanTitle, "--") > 0
        cleanTitle = Replace(cleanTitle, "--", "-")
    Loop
    
    ' Check if numbering contains dots (sub-chapter) or not (major chapter)
    Dim anchorResult As String
    If InStr(numbering, ".") > 0 Then
        ' Sub-chapter: use HTML anchor format (ch + numbers without dots)
        anchorResult = "ch" & Replace(numbering, ".", "-")
        If cleanTitle <> "" Then anchorResult = anchorResult & "-" & cleanTitle
    Else
        ' Major chapter: use standard markdown anchor
        anchorResult = numbering
        If cleanTitle <> "" Then anchorResult = anchorResult & "-" & cleanTitle
    End If
    GenerateTOCAnchor = anchorResult
End Function

' Generate markdown heading from document heading node
' @param node: Document node containing heading information
' @return: Formatted markdown heading string
Function Head(ByRef node)
    Dim i : i = Split(node.name_, " ")(1)  ' Extract heading level from node name
    Dim headingText As String : headingText = ""
    
    ' Safely extract heading text with error handling
    On Error Resume Next
    If node.children.Count > 0 Then
        ' Try different methods to get the text
        headingText = node.children(1).value.getString()
        If headingText = "" Then
            headingText = node.children(1).value.String
        End If
    End If
    On Error GoTo 0
    
    ' If still empty, use fallback
    If headingText = "" Then
        headingText = "Heading " & i
    End If
    
    ' Generate heading with HTML anchor for sub-chapters (hybrid approach)
    Dim headingOutput As String
    ' Extract numbering part to check for dots
    Dim spacePos As Integer : spacePos = InStr(headingText, " ")
    Dim numbering As String : numbering = ""
    If spacePos > 0 Then numbering = Left(headingText, spacePos - 1)
    
    If InStr(numbering, ".") > 0 Then
        ' Sub-chapter: add HTML anchor
        Dim anchor As String : anchor = GenerateTOCAnchor(headingText)
        headingOutput = CHR$(10) & "# <a id=""" & anchor & """></a>" & headingText & "  " & CHR$(10)
    Else
        ' Major chapter: use simple markdown heading
        headingOutput = CHR$(10) & String(i, "#") & " " & headingText & "  " & CHR$(10)
    End If
    
    Head = headingOutput
End Function

' Generate markdown blockquote from document node
' @param node: Document node containing quote content
' @return: Formatted markdown blockquote string
Function Quote(ByRef node)
    Quote = "> " & docView.PrintTree(node) & "  " & CHR$(10)
End Function

' Generate markdown code block from document node
' @param node: Document node containing code content
' @return: Formatted markdown code block with language specification
Function Code(ByRef node)
    Dim codeContent As String  ' Variable to store processed code content
    If docView.props("CodeLineNum") Then
        Dim props As New Collection
        props.Add(True, "CodeLineNum")
        codeContent = docView.PrintTree(node, props)
    Else
        codeContent = docView.PrintTree(node)
    End If
    ' Extract language from node name and format as code block
    Code = "```" & Split(node.name_, "_")(1) & CHR$(10) & _
        codeContent & "```" & "  " & CHR$(10)
End Function

' Process paragraph style node and return formatted content
' @param node: Document node with paragraph style information
' @return: Formatted paragraph content
Function ParaStyle(ByRef node)
    ParaStyle = docView.PrintTree(node)
End Function

' Process list node and return formatted markdown list
' @param node: Document node containing list structure
' @return: Formatted markdown list content
Function List(ByRef node)
    List = docView.PrintTree(node)
End Function

' Generate markdown image syntax from LibreOffice image object
' @param lo: LibreOffice image object
' @return: Formatted markdown image string with alt text and description
Function Image(ByRef lo)
    Dim imageUrl As String : imageUrl = docView.ProcessImage(lo, docView.docURL)
    Image = imageUrl & "  " & CHR$(10)
End Function

' Generate inline HTML image tag from LibreOffice image object
' @param lo: LibreOffice image object
' @return: HTML img tag with inline attribute
Function InlineImage(ByRef lo)
    Dim imageUrl As String : imageUrl = docView.ProcessImage(lo, docView.docURL)
    ' Extract src from markdown format
    Dim srcStart As Long : srcStart = InStr(imageUrl, "](")
    Dim srcEnd As Long : srcEnd = InStr(srcStart, imageUrl, ")")
    Dim src As String : src = Mid(imageUrl, srcStart + 2, srcEnd - srcStart - 2)
    ' Remove ./ prefix if present
    If Left(src, 2) = "./" Then src = Mid(src, 3)
    InlineImage = "<img inline=""true"" src=""" & src & """ />"
End Function

' Generate markdown link from document hyperlink node
' @param node: Document node containing hyperlink information
' @return: Formatted markdown link string
Function Link(ByRef node)
    Dim t As String : t = node.String  ' Extract link text
    Dim url As String : url = node.HyperLinkURL
    Dim linkResult As String
    
    ' Remove trailing characters for heading styles
    If Left(node.ParaStyleName, 8) = STYLE_HEADING Then
        t = Left(t, Len(t) - 2)
    End If
    
    ' Convert LibreOffice internal references to GitHub-style markdown anchors
    If InStr(url, "#") > 0 Then
        Dim anchor As String : anchor = Mid(url, InStr(url, "#") + 1)
        ' Clean up LibreOffice anchor names for GitHub markdown
        If InStr(anchor, "__RefHeading") > 0 Then
            ' Hybrid approach for TOC links
            anchor = GenerateTOCAnchor(t)
        End If
        linkResult = "[" & t & "](#" & anchor & ")"
    Else
        linkResult = "[" & t & "](" & url & ")"
    End If
    
    ' Format Table of Contents links without indentation
    If Left(node.ParaStyleName, 8) = STYLE_HEADING Then
        Link = linkResult & "  " & CHR$(10) & "  " & CHR$(10)
    Else
        Link = linkResult
    End If
End Function

' Generate HTML anchor tag from LibreOffice bookmark object
' @param lo: LibreOffice bookmark object
' @return: HTML anchor tag or empty string
Function Anchor(ByRef lo)
    If lo.IsStart Then
        Dim anchorName As String : anchorName = lo.Bookmark.Name
        ' Clean up LibreOffice bookmark names for markdown
        If InStr(anchorName, "__RefHeading") > 0 Then
            ' Skip LibreOffice internal bookmarks - they'll be handled by headings
            Anchor = ""
        Else
            Anchor = "<a id=""" & anchorName & """></a>"
        End If
    Else
        Anchor = ""
    End If
End Function

' Apply font decoration (bold, italic, etc.) to text node
' @param node: Text node to decorate
' @param style: Style name (Bold, Italic, Underline, Strikeout)
' @return: Text wrapped with appropriate markdown/HTML tags
Function FontDecorate(ByRef node, style As String)
    Dim s : s = fontStyles(style)  ' Get decoration markers for the style
    FontDecorate = s(0) & node.String & s(1)
End Function

' Format table cell content for markdown table
' @param txt: Cell content text
' @param level: Nesting level (unused)
' @param index: Column index in the row
' @param idxRow: Row index (unused)
' @return: Formatted cell content with pipe separator
Function FormatCell(ByRef txt, level As Long, index As Long, idxRow As Long)
    ' Clean up cell content: remove line breaks and trim whitespace
    Dim cleanTxt As String
    cleanTxt = Replace(txt, CHR$(10), " ")
    cleanTxt = Replace(cleanTxt, CHR$(13), " ")
    cleanTxt = Trim(cleanTxt)
    
    ' Always add pipe separator before cell content
    FormatCell = "|" & cleanTxt
End Function

' Format table row for markdown table with header separator
' @param txt: Row content text
' @param level: Nesting level (unused)
' @param index: Row index (0 for header row)
' @param Colls: Number of columns
' @return: Formatted markdown table row with separator after header
Function FormatRow(ByRef txt, level As Long, index As Long, Colls As Long)
    Dim i AS Long, r As String : r = ""  ' Loop counter and result string
    r = txt & "|" & CHR$(10)
    ' Add markdown table header separator after first row
    If index = 0 Then
        r = r & "|"
        For i = 0 To Colls
            r = r & " --- |"
        Next
        r = r & CHR$(10)
    End If
    FormatRow = r
End Function

' Format complete table content
' @param txt: Complete table content
' @param level: Nesting level (unused)
' @return: Formatted table content (pass-through)
Function FormatTable(ByRef txt, level As Long)
    FormatTable = txt
End Function

' Format list item with proper indentation and markers
' @param list: LibreOffice list object containing numbering information
' @param txt: List item text content
' @param level: Nesting level (unused)
' @return: Formatted markdown list item with indentation
Function FormatList(ByRef list, ByRef txt, level As Long)
    Dim shift As String : shift = String(list.NumberingLevel * SHIFT_CNT, " ")  ' Calculate indentation
    Dim lbl As String : lbl = list.ListLabelString  ' Get list marker (number or bullet)
    ' Clean up text and ensure proper line breaks
    Dim cleanTxt As String : cleanTxt = Trim(txt)
    FormatList = shift & IIf(lbl = "", "-", lbl) & " " & cleanTxt & "  " & CHR$(10)
End Function

' Format paragraph with optional extra line break
' @param txt: Paragraph text content
' @param level: Nesting level (unused)
' @param extra: Flag to add extra line break (0 = no, other = yes)
' @return: Formatted paragraph text
Function FormatPara(ByRef txt, level As Long, extra As Long)
    ' Ensure proper line breaks for markdown formatting
    ' Add double spaces before line break for hard line breaks in Markdown
    FormatPara = txt & "  " & CHR$(10)
End Function

' Convert LibreOffice formula to LaTeX format for markdown
' @param txt: Formula text from LibreOffice
' @return: LaTeX formula wrapped in markdown math delimiters
Function Formula(ByRef txt As String)
    Dim m As New mMath          ' Math formula processor
    m.Set_Formula(txt)          ' Set the input formula
    m.vAdapter = New vLatex     ' Use LaTeX output adapter
    m.vAdapter.mMath = m        ' Link adapter to math processor
    Dim formulaContent As String : formulaContent = m.Get_Formula()
    Formula = "$$" & "  " & CHR$(10) & formulaContent & "  " & CHR$(10) &  "$$" & "  " & CHR$(10)
End Function

' Extract and remove section title from section node
' @param nodeSec: Section node to extract title from
' @return: Section title string or empty string if no title found
Function GetSectionTitle(ByRef nodeSec)
    GetSectionTitle = ""  ' Default return value
    ' Check if section has style children
    If nodeSec.children.Count > 0 And _
        nodeSec.children(1).type_ =  NodeType.Style Then
        Dim nodeStyle : nodeStyle = nodeSec.children(1)
        ' Check if style has paragraph children
        If nodeStyle.children.Count > 0 And _
            nodeStyle.children(1).type_ =  NodeType.Paragraph Then
            GetSectionTitle = nodeStyle.children(1).value.getString()
            nodeStyle.children.Remove(1)  ' Remove title paragraph after extraction
        End If
    End If
End Function

' Format document section with appropriate markdown/HTML structure
' @param nodeSec: Section node to format
' @return: Formatted section content (heading or spoiler block)
Function Section(ByRef nodeSec)
    ' For non-top-level sections, just print content
    If nodeSec.level <> 1 Then
        Section = docView.PrintTree(nodeSec)
        Exit Function
    End If
    
    Dim secTitle : secTitle = GetSectionTitle(nodeSec)  ' Extract section title
    ' Handle main heading sections
    If secTitle = SEC_HEADING Then
        Section = "# " & secTitle & CHR$(10) & docView.PrintTree(nodeSec)
        Exit Function
    End If

    ' Check if this is Table of Contents section
    Dim lowerTitle As String : lowerTitle = LCase(secTitle)
    If InStr(lowerTitle, "оглавление") > 0 Or InStr(lowerTitle, "содержание") > 0 Or _
       InStr(lowerTitle, "индекс") > 0 Or InStr(lowerTitle, "список глав") > 0 Or _
       InStr(lowerTitle, "указатель") > 0 Or InStr(lowerTitle, "каталог") > 0 Or _
       InStr(lowerTitle, "реестр") > 0 Or InStr(lowerTitle, "contents") > 0 Or _
       InStr(lowerTitle, "index") > 0 Or InStr(lowerTitle, "table of contents") > 0 Or _
       InStr(lowerTitle, "reference") > 0 Then
        ' Use regular heading for TOC
        Section = "## " & secTitle & "  " & CHR$(10) & "  " & CHR$(10) & docView.PrintTree(nodeSec)
    Else
        ' Wrap other sections in spoiler tags for HFM
        Section = "<spoiler title=""" & _
            secTitle & """>" & "  " & CHR$(10) & "  " & CHR$(10) & _
            docView.PrintTree(nodeSec) & "</spoiler>" & "  " & CHR$(10)
    End If
End Function
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

' Generate markdown heading from document heading node
' @param node: Document node containing heading information
' @return: Formatted markdown heading string
Function Head(ByRef node)
    Dim i : i = Split(node.name_, " ")(1)  ' Extract heading level from node name
    Dim headingText As String : headingText = node.children(1).value.String
    
    ' Create simple markdown heading (GitHub will auto-generate anchors)
    Head = CHR$(10) & String(i, "#") & " " & headingText & CHR$(10)
End Function

' Generate markdown blockquote from document node
' @param node: Document node containing quote content
' @return: Formatted markdown blockquote string
Function Quote(ByRef node)
    Quote = "> " & docView.PrintTree(node) & CHR$(10)
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
        codeContent & "```" & CHR$(10)
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
    Image = "![" & lo.Title & "](" & lo.Graphic.OriginURL & _
        " """ & lo.Description & """)" & CHR$(10)
End Function

' Generate inline HTML image tag from LibreOffice image object
' @param lo: LibreOffice image object
' @return: HTML img tag with inline attribute
Function InlineImage(ByRef lo)
    InlineImage = "<img inline=""true"" src=""" & _
        lo.Graphic.OriginURL & """ />"
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
            ' Convert to GitHub-style anchor based on text content
            ' GitHub converts: "2.1. Ссылки" -> "21-ссылки"
            anchor = LCase(t)
            anchor = Replace(anchor, ".", "")
            anchor = Replace(anchor, " ", "-")
            anchor = Replace(anchor, ",", "")
            anchor = Replace(anchor, "(", "")
            anchor = Replace(anchor, ")", "")
        End If
        linkResult = "[" & t & "](#" & anchor & ")"
    Else
        linkResult = "[" & t & "](" & url & ")"
    End If
    
    ' Format Table of Contents links as list items
    If Left(node.ParaStyleName, 8) = STYLE_HEADING Then
        ' Count dots to determine nesting level
        Dim dotCount As Long : dotCount = Len(t) - Len(Replace(t, ".", ""))
        Dim indent As String : indent = String((dotCount - 1) * 2, " ")
        Link = indent & "- " & linkResult & CHR$(10)
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
    FormatList = shift & IIf(lbl = "", "-", lbl) & " " & cleanTxt & CHR$(10)
End Function

' Format paragraph with optional extra line break
' @param txt: Paragraph text content
' @param level: Nesting level (unused)
' @param extra: Flag to add extra line break (0 = no, other = yes)
' @return: Formatted paragraph text
Function FormatPara(ByRef txt, level As Long, extra As Long)
    ' Ensure proper line breaks for markdown formatting
    FormatPara = txt & CHR$(10)
End Function

' Convert LibreOffice formula to LaTeX format for markdown
' @param txt: Formula text from LibreOffice
' @return: LaTeX formula wrapped in markdown math delimiters
Function Formula(ByRef txt As String)
    Dim m As New mMath          ' Math formula processor
    m.Set_Formula(txt)          ' Set the input formula
    m.vAdapter = New vLatex     ' Use LaTeX output adapter
    m.vAdapter.mMath = m        ' Link adapter to math processor
    Formula = "$$" & CHR$(10) & m.Get_Formula() & CHR$(10) &  "$$" & CHR$(10)
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
        Section = "## " & secTitle & CHR$(10) & CHR$(10) & docView.PrintTree(nodeSec)
    Else
        ' Wrap other sections in spoiler tags for HFM
        Section = "<spoiler title=""" & _
            secTitle & """>" & CHR$(10) & CHR$(10) & _
            docView.PrintTree(nodeSec) & "</spoiler>" & CHR$(10)
    End If
End Function
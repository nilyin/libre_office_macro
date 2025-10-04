REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option VBASupport 1

' Constants for document processing
Const STYLE_HEAD = "Heading"  ' Prefix for heading paragraph styles
Const CODE_LINE_NUM = True    ' Enable line numbering in code blocks

' File logging function for headless mode debugging
Sub LogToFile(message As String)
    Dim fileNum As Integer : fileNum = FreeFile
    Open "/tmp/macro_debug.log" For Append As #fileNum
    Print #fileNum, Now() & " - " & message
    Close #fileNum
End Sub

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
                (Len(curPara.ParaStyleName) < 7 Or Left(curPara.ParaStyleName, 7) <> STYLE_HEAD) Then
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
    Dim docPath As String : docPath = convertFromURL(Comp.URL)
    ' Handle both .odt and .fodt extensions
    If Right(LCase(docPath), 5) = ".fodt" Then
        docPath = Left(docPath, Len(docPath) - 5) & suffix
    ElseIf Right(LCase(docPath), 4) = ".odt" Then
        docPath = Left(docPath, Len(docPath) - 4) & suffix
    Else
        docPath = docPath & suffix
    End If
    Filename = convertToURL(docPath)
	FileNo = Freefile
	Open Filename For Output As #FileNo
	Print #FileNo, text_
End Sub

Function ProcessHeader(ByRef Comp As Object) As String
    Dim headerContent As String : headerContent = ""
    On Error Resume Next
    
    ' Try to get the current page style from the document
    Dim cursor : cursor = Comp.getText().createTextCursor()
    Dim pageStyleName As String : pageStyleName = cursor.PageStyleName
    
    ' Get page styles collection
    Dim pageStyles : pageStyles = Comp.getStyleFamilies().getByName("PageStyles")
    Dim currentPageStyle : currentPageStyle = pageStyles.getByName(pageStyleName)
    
    ' Process header if it exists
    If currentPageStyle.HeaderIsOn Then
        Dim headerText : headerText = currentPageStyle.HeaderText
        If Not IsEmpty(headerText) Then
            Dim headerEnum : headerEnum = headerText.createEnumeration()
            Do While headerEnum.hasMoreElements()
                Dim headerElement : headerElement = headerEnum.nextElement()
                If headerElement.supportsService("com.sun.star.text.Paragraph") Then
                    ' Process text portions first to catch Frame type images
                    Dim textEnum : textEnum = headerElement.createEnumeration()
                    Do While textEnum.hasMoreElements()
                        Dim textPortion : textPortion = textEnum.nextElement()
                        If textPortion.TextPortionType = "Text" Then
                            headerContent = headerContent & textPortion.String
                        ElseIf textPortion.TextPortionType = "Frame" Then
                            ' Process inline frames (images)
                            Dim frameEnum : frameEnum = textPortion.createContentEnumeration("com.sun.star.text.TextGraphicObject")
                            Do While frameEnum.hasMoreElements()
                                Dim imageObj : imageObj = frameEnum.nextElement()
                                If imageObj.supportsService("com.sun.star.text.TextGraphicObject") Then
                                    headerContent = headerContent & ProcessHeaderImage(imageObj, Comp.URL)
                                End If
                            Loop
                        End If
                    Loop
                    
                    ' Also check for paragraph-anchored content
                    Dim contentEnum : contentEnum = headerElement.createContentEnumeration("com.sun.star.text.TextContent")
                    Do While contentEnum.hasMoreElements()
                        Dim content : content = contentEnum.nextElement()
                        If content.supportsService("com.sun.star.text.TextGraphicObject") Then
                            headerContent = headerContent & ProcessHeaderImage(content, Comp.URL)
                        End If
                    Loop
                    
                    headerContent = headerContent & "  " & CHR$(10)
                End If
            Loop
            If headerContent <> "" Then headerContent = headerContent & "  " & CHR$(10)
        End If
    End If
    
    On Error GoTo 0
    ProcessHeader = headerContent
End Function



' Generate image folder name from document URL with robust extraction
' @param docURL: Document URL
' @return: Image folder name with pattern "img_" + source filename
Private Function GenerateImageFolderName(ByRef docURL As String) As String
    On Error Resume Next
    
    ' Enhanced filename extraction with multiple fallback methods
    Dim fileName As String
    
    ' Method 1: Standard ConvertFromURL
    fileName = ConvertFromURL(docURL)
    If fileName <> "" Then
        fileName = Mid(fileName, InStrRev(fileName, GetPathSeparator()) + 1)
    End If
    
    ' Method 2: Direct URL parsing if ConvertFromURL fails
    If fileName = "" Then
        fileName = docURL
        If InStr(fileName, "/") > 0 Then
            fileName = Mid(fileName, InStrRev(fileName, "/") + 1)
        End If
        If InStr(fileName, "\") > 0 Then
            fileName = Mid(fileName, InStrRev(fileName, "\") + 1)
        End If
    End If
    
    ' Method 3: Extract from file:// URL format
    If fileName = "" And Left(docURL, 7) = "file://" Then
        fileName = Mid(docURL, 8)
        fileName = Replace(fileName, "/", GetPathSeparator())
        fileName = Mid(fileName, InStrRev(fileName, GetPathSeparator()) + 1)
    End If
    
    ' Remove extension and create folder name
    Dim dotPos As Long : dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        fileName = Left(fileName, dotPos - 1)
    End If
    
    ' Fallback to generic name if all methods fail
    If fileName = "" Then fileName = "document"
    
    GenerateImageFolderName = "img_" & fileName
    On Error GoTo 0
End Function

Function ProcessHeaderImage(ByRef imageObj, ByRef docURL As String) As String
    Dim altText As String : altText = IIf(imageObj.Title = "", "logo", imageObj.Title)
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
            ProcessHeaderImage = "![" & altText & "](" & imageObj.HyperLinkURL & ")" & "  " & CHR$(10) & "  " & CHR$(10)
        End If
    End If
    On Error GoTo 0
    
    If hasExternalLink Then
        Exit Function
    End If
    
    ' Check if it's a remote URL in the image source itself
    If imageName <> "" And Left(LCase(imageName), 4) = "http" Then
        ProcessHeaderImage = "![" & altText & "](" & imageName & ")" & "  " & CHR$(10) & "  " & CHR$(10)
        Exit Function
    End If
    
    ' For header images, use a simple naming scheme (only extract first header image)
    Dim fileName As String : fileName = "header-logo.png"
    Dim docDir As String : docDir = Left(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), GetPathSeparator()))
    
    ' Try to extract the header image
    On Error Resume Next
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    Dim imgFolderName As String : imgFolderName = GenerateImageFolderName(docURL)
    Dim imgDir As String : imgDir = docDir & imgFolderName
    Dim targetPath As String : targetPath = imgDir & GetPathSeparator() & fileName
    
    ' Create img directory if it doesn't exist
    If Not fso.FolderExists(imgDir) Then fso.CreateFolder(imgDir)
    
    ' Try to extract using GraphicProvider
    Dim graphic : graphic = imageObj.Graphic
    If Not IsEmpty(graphic) Then
        Dim graphicProvider : graphicProvider = CreateUnoService("com.sun.star.graphic.GraphicProvider")
        Dim exportProps(1) As New com.sun.star.beans.PropertyValue
        exportProps(0).Name = "URL"
        exportProps(0).Value = ConvertToURL(targetPath)
        exportProps(1).Name = "MimeType"
        exportProps(1).Value = "image/png"
        graphicProvider.storeGraphic(graphic, exportProps())
    End If
    On Error GoTo 0
    
    ProcessHeaderImage = "![" & altText & "](" & imgFolderName & "/" & fileName & ")" & "  " & CHR$(10) & "  " & CHR$(10)
End Function

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
    dView.imageCounter = 0
    dView.docPrefix = ""
    dView.docURL = doc.URL
    dView.props = New Collection
    With dView.props
        .Add(CODE_LINE_NUM, "CodeLineNum")
    End With
    Dim headerContent As String : headerContent = ProcessHeader(doc)
    headerContent = vHtml.ConvertMarkdownImages(headerContent)
    Dim fullContent As String : fullContent = headerContent & dView.MakeView()
    ExportToFile fullContent, doc, ".html"
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
    dView.imageCounter = 0
    dView.docPrefix = ""
    dView.docURL = doc.URL
    dView.props = New Collection
    With dView.props
        .Add(CODE_LINE_NUM, "CodeLineNum")
    End With
    Dim headerContent As String : headerContent = ProcessHeader(doc)
    Dim fullContent As String : fullContent = headerContent & dView.MakeView()
    ExportToFile fullContent, doc, ".md"
End Sub

' "C:\Program Files\LibreOffice\program\soffice.exe" --invisible --nofirststartwizard --headless --norestore macro:///DocExport.DocModel.ExportDir("D:\cpp\habr\002-hfm",0)
Sub ExportDir(Folder As String, Optional Hfm As Variant)
    Dim useHfm As Boolean : useHfm = True
    If Not IsMissing(Hfm) Then
        If IsNumeric(Hfm) Then useHfm = CBool(Hfm)
        If VarType(Hfm) = vbBoolean Then useHfm = Hfm
    End If
    
    ' Enhanced logging
    Dim logFile As String : logFile = Folder & GetPathSeparator() & "conversion.log"
    LogDebug "=== LibreOffice Conversion Started ===", logFile
    LogDebug "Folder: " & Folder, logFile
    LogDebug "HFM Mode: " & useHfm, logFile
    LogDebug "Path Separator: " & GetPathSeparator(), logFile
    
    ' Get ODT files using robust method
    Dim fileCount As Long : fileCount = GetODTFiles(Folder)
    LogDebug "ODT Files Found: " & fileCount, logFile
    
    If fileCount > 0 Then
        ' Get actual file list
        Dim odtFiles : odtFiles = GetODTFileList(Folder)
        Dim Props(0) As New com.sun.star.beans.PropertyValue
        Props(0).Name = "Hidden"
        Props(0).Value = True
        
        Dim Comp As Object
        Dim i As Long
        For i = 0 To fileCount - 1
            Dim filePath As String : filePath = odtFiles(i)
            LogDebug "Processing: " & filePath, logFile
            
            Dim url As String : url = ConvertToURL(filePath)
            LogDebug "URL: " & url, logFile
            
            On Error Resume Next
            Comp = StarDesktop.loadComponentFromURL(url, "_blank", 0, Props)
            If Err.Number = 0 And Not IsEmpty(Comp) Then
                LogDebug "Document loaded successfully", logFile
                
                If useHfm Then
                    MakeDocHfmView Comp
                    LogDebug "HFM conversion completed", logFile
                Else
                    MakeDocHtmlView Comp
                    LogDebug "HTML conversion completed", logFile
                End If
                
                Comp.close(True)
                LogDebug "Document closed", logFile
            Else
                LogDebug "ERROR: Failed to load document - " & Err.Description, logFile
            End If
            On Error GoTo 0
        Next
    Else
        LogDebug "ERROR: No ODT files found", logFile
    End If
    
    LogDebug "=== Conversion Complete ===", logFile
End Sub

' Headless wrapper function for command line execution
' Based on SuperUser solution pattern
Sub HeadlessBatch(FolderPath As String, Optional UseHfm As Variant)
    Dim useHfmMode As Boolean : useHfmMode = True
    If Not IsMissing(UseHfm) Then
        If IsNumeric(UseHfm) Then useHfmMode = CBool(UseHfm)
        If VarType(UseHfm) = vbBoolean Then useHfmMode = UseHfm
    End If
    
    LogToFile "HeadlessBatch: Starting for " & FolderPath
    
    ' Get file count first
    Dim fileCount As Long : fileCount = GetODTFiles(FolderPath)
    LogToFile "HeadlessBatch: Found " & fileCount & " files"
    
    If fileCount > 0 Then
        ' Get actual file list
        Dim odtFiles : odtFiles = GetODTFileList(FolderPath)
        
        Dim Props(0) As New com.sun.star.beans.PropertyValue
        Props(0).Name = "Hidden"
        Props(0).Value = True
        
        Dim i As Long
        For i = 0 To fileCount - 1
            Dim filePath As String : filePath = odtFiles(i)
            LogToFile "HeadlessBatch: Processing " & filePath
            
            Dim url As String : url = ConvertToURL(filePath)
            Dim Document : Document = StarDesktop.loadComponentFromURL(url, "_blank", 0, Props)
            
            If Not IsEmpty(Document) Then
                If useHfmMode Then
                    MakeDocHfmView Document
                Else
                    MakeDocHtmlView Document
                End If
                Document.close(True)
                LogToFile "HeadlessBatch: Completed " & filePath
            Else
                LogToFile "HeadlessBatch: Failed to load " & filePath
            End If
        Next
    End If
    
    LogToFile "HeadlessBatch: Finished"
End Sub

' Robust file enumeration with multiple fallback methods
Function GetODTFiles(ByRef folderPath As String) As Variant
    Dim fileList() As String
    Dim fileCount As Long : fileCount = 0
    
    LogDebug "GetODTFiles: Starting enumeration for " & folderPath
    
    On Error Resume Next
    
    ' Method 1: Try Dir$ function for .odt files
    Dim fileName As String : fileName = Dir$(folderPath & GetPathSeparator() & "*.odt", 0)
    LogDebug "GetODTFiles: Dir$ first .odt file = " & fileName
    Do While fileName <> ""
        ReDim Preserve fileList(fileCount)
        fileList(fileCount) = folderPath & GetPathSeparator() & fileName
        LogDebug "GetODTFiles: Added file " & fileList(fileCount)
        fileCount = fileCount + 1
        fileName = Dir$
    Loop
    
    ' Also search for .fodt files
    fileName = Dir$(folderPath & GetPathSeparator() & "*.fodt", 0)
    LogDebug "GetODTFiles: Dir$ first .fodt file = " & fileName
    Do While fileName <> ""
        ReDim Preserve fileList(fileCount)
        fileList(fileCount) = folderPath & GetPathSeparator() & fileName
        LogDebug "GetODTFiles: Added file " & fileList(fileCount)
        fileCount = fileCount + 1
        fileName = Dir$
    Loop
    LogDebug "GetODTFiles: Dir$ method found " & fileCount & " files"
    
    ' Method 2: If Dir$ failed, try LibreOffice SimpleFileAccess
    If fileCount = 0 Then
        LogDebug "GetODTFiles: Trying SimpleFileAccess method"
        Dim fileAccess : fileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
        If Not IsEmpty(fileAccess) Then
            Dim urlPath As String : urlPath = ConvertToURL(folderPath)
            LogDebug "GetODTFiles: URL path = " & urlPath
            If fileAccess.exists(urlPath) Then
                Dim contents : contents = fileAccess.getFolderContents(urlPath, False)
                LogDebug "GetODTFiles: Folder contents count = " & (UBound(contents) + 1)
                Dim i As Long
                For i = 0 To UBound(contents)
                    Dim fileUrl As String : fileUrl = contents(i)
                    If Right(LCase(fileUrl), 4) = ".odt" Or Right(LCase(fileUrl), 5) = ".fodt" Then
                        ReDim Preserve fileList(fileCount)
                        fileList(fileCount) = ConvertFromURL(fileUrl)
                        LogDebug "GetODTFiles: SimpleFileAccess added " & fileList(fileCount)
                        fileCount = fileCount + 1
                    End If
                Next
            Else
                LogDebug "GetODTFiles: Folder does not exist via SimpleFileAccess"
            End If
        Else
            LogDebug "GetODTFiles: SimpleFileAccess service not available"
        End If
    End If
    
    On Error GoTo 0
    
    LogDebug "GetODTFiles: Total files found = " & fileCount
    GetODTFiles = fileCount
End Function

' Get actual ODT file list array
Function GetODTFileList(ByRef folderPath As String) As Variant
    Dim fileList() As String
    Dim fileCount As Long : fileCount = 0
    
    On Error Resume Next
    
    ' Method 1: Try Dir$ function for .odt files
    Dim fileName As String : fileName = Dir$(folderPath & GetPathSeparator() & "*.odt", 0)
    Do While fileName <> ""
        ReDim Preserve fileList(fileCount)
        fileList(fileCount) = folderPath & GetPathSeparator() & fileName
        fileCount = fileCount + 1
        fileName = Dir$
    Loop
    
    ' Also search for .fodt files
    fileName = Dir$(folderPath & GetPathSeparator() & "*.fodt", 0)
    Do While fileName <> ""
        ReDim Preserve fileList(fileCount)
        fileList(fileCount) = folderPath & GetPathSeparator() & fileName
        fileCount = fileCount + 1
        fileName = Dir$
    Loop
    
    ' Method 2: If Dir$ failed, try LibreOffice SimpleFileAccess
    If fileCount = 0 Then
        Dim fileAccess : fileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
        If Not IsEmpty(fileAccess) Then
            Dim urlPath As String : urlPath = ConvertToURL(folderPath)
            If fileAccess.exists(urlPath) Then
                Dim contents : contents = fileAccess.getFolderContents(urlPath, False)
                Dim i As Long
                For i = 0 To UBound(contents)
                    Dim fileUrl As String : fileUrl = contents(i)
                    If Right(LCase(fileUrl), 4) = ".odt" Or Right(LCase(fileUrl), 5) = ".fodt" Then
                        ReDim Preserve fileList(fileCount)
                        fileList(fileCount) = ConvertFromURL(fileUrl)
                        fileCount = fileCount + 1
                    End If
                Next
            End If
        End If
    End If
    
    On Error GoTo 0
    GetODTFileList = fileList
End Function


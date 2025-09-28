# LibreOffice Export Macros - Working Documentation

## Overview
The export macros convert LibreOffice Writer documents (.odt) to Habr Flavored Markdown (HFM) or HTML by parsing the document structure and applying format-specific transformations.

## Step-by-Step Export Process

### A. Macro Installation via Extension (Recommended)

To make the export macros available globally across all LibreOffice documents, they can be packaged into an `.oxt` extension file and installed.

#### 1. Creating the Extension File
A Python script, `create_oxt.py`, is provided to automate this process.

**To run the script:**
```bash
python create_oxt.py
```
This command reads all the necessary `.bas` files, generates the required XML metadata (`description.xml`, `manifest.xml`, etc.), and packages everything into a single `DocExport.oxt` file in the project's root directory.

#### 2. Installing the Extension from Command Line
Once the `DocExport.oxt` file is created, it can be installed into LibreOffice using the `unopkg` command-line tool. Using the `--shared` flag installs the extension for all users, making the macros globally accessible.

**Windows:**
```bash
"C:\Program Files\LibreOffice\program\unopkg.com" add --shared "C:\path\to\your\project\DocExport.oxt"
```

**macOS:**
```bash
/Applications/LibreOffice.app/Contents/MacOS/unopkg add --shared /path/to/your/project/DocExport.oxt
```

**Linux:**
```bash
/usr/lib/libreoffice/program/unopkg add --shared /path/to/your/project/DocExport.oxt
```

After installation, the `MakeDocHfmView` and `MakeDocHtmlView` macros will be available in any LibreOffice Writer document via `Tools -> Macros -> Run Macro...` under the `DocExport` library.

### B. Manual Export Process (Per-Document Macros)

This process describes running the macros embedded within a single document. This is useful for development but less convenient for general use. The `update_macros.py` script is used to keep the embedded macros in sync with the source `.bas` files.

**To update embedded macros:**
```bash
python update_macros.py
```

### 1. Document Parsing (`MakeModel` function in DocModel.bas)
```
Document (.odt) → Document Tree Structure → Formatted Output
```

**Steps:**
1. **Initialize**: Create root document tree node with `NodeType.Section`
2. **Enumerate**: Get paragraph enumerator from `Comp.getText().createEnumeration()`
3. **Parse**: Call `SectionParse()` to recursively build document tree
4. **Return**: Complete document tree structure

### 2. Tree Building Process (`SectionParse` function)

**For each document element:**
- **Tables**: Create `NodeType.Table` nodes with LibreOffice table object
- **Paragraphs**: Process through style and section hierarchy
- **Sections**: Handle document sections (nested content blocks)
- **Lists**: Group numbered/bulleted items under `NodeType.List` nodes

### 3. Document Tree Structure

**Global Variable**: `docTree` (in DocView class)
- **Type**: `Node` structure (defined in DocModel.bas)
- **Format**: Hierarchical tree with 5 node types:

```
Node Structure:
├── type_: NodeType (Section|Style|List|Table|Paragraph)
├── value: LibreOffice object (for Table/Paragraph nodes)
├── name_: String identifier (for Section/Style nodes)
├── children: Collection of child nodes
└── level: Integer nesting depth
```

**Memory Storage Example:**
```
Root Section (level=0)
├── Style "Heading 1" (level=1)           [Document Order: 1st]
│   └── Paragraph (level=2, value=LO_Paragraph_Object)
├── Style "code_cpp" (level=1)            [Document Order: 2nd]
│   ├── Paragraph (level=2, value=LO_Paragraph_Object)
│   └── Paragraph (level=2, value=LO_Paragraph_Object)
└── Table (level=1, value=LO_Table_Object) [Document Order: 3rd]
```

### Node Ordering and Output Generation

**Critical**: Order in the `children` Collection determines final output sequence.

**How Ordering Works:**
1. **Document Parsing**: Elements are added to `children` Collection in document appearance order
2. **Collection Iteration**: `For Each child In node.children` preserves insertion order
3. **Output Generation**: PrintTree processes nodes sequentially, maintaining document flow

**Example with Same-Level Elements:**
```
Style "Normal" (level=1)
├── Paragraph "Introduction text" (level=2)     [Output: 1st]
├── Table "Data table" (level=2)               [Output: 2nd] 
└── Paragraph "Conclusion text" (level=2)      [Output: 3rd]
```

**Result**: Text → Table → Text (exactly as they appear in the document)

**Key Point**: LibreOffice Basic Collections maintain insertion order, ensuring document structure is preserved during export.

## PrintTree Function Purpose

**Location**: `DocView.bas`
**Purpose**: Recursive tree traversal and content generation with order preservation

### Function Logic:
1. **Initialize**: Set line numbering based on `props` collection
2. **Traverse**: Iterate through all child nodes **in document order**
3. **Dispatch**: Route each node type to appropriate formatter:
   - `NodeType.Section` → `viewAdapter.Section()`
   - `NodeType.Style` → `PrintNodeStyle()`
   - `NodeType.List` → `viewAdapter.List()`
   - `NodeType.Paragraph` → `PrintNodePara()`
   - `NodeType.Table` → `PrintNodeTable()`
4. **Accumulate**: Concatenate formatted output strings **preserving order**
5. **Return**: Complete formatted document content

### Order Preservation Mechanism:
```basic
' Critical code in PrintTree:
For Each child In node.children  ' Iterates in insertion order
    If child.type_ = NodeType.Section Then
        s = s & viewAdapter.Section(child)     ' Append in sequence
    ElseIf child.type_ = NodeType.Table Then
        s = s & PrintNodeTable(child)         ' Maintain document flow
    ' ... other node types
Next
```

**Result**: Code blocks and tables appear in final output exactly as positioned in original document.

### Key Parameters:
- `node`: Current tree node being processed
- `props`: Optional Collection with formatting flags (detailed below)

## Props Object - Detailed Structure

### Object Type and Purpose
**Type**: LibreOffice Basic `Collection` object
**Purpose**: Runtime configuration for formatting options during document processing
**Scope**: Passed down through recursive PrintTree calls

### Structure and Methods
```basic
' Collection Methods Used:
props.Add(value, key)     ' Add configuration option
props(key)               ' Retrieve configuration value
props.Count              ' Number of configuration items
```

### Current Implementation Usage
**Location**: Created in `MakeDocHfmView`/`MakeDocHtmlView` functions (DocModel.bas)

**Example Content for Parsed Document:**
```basic
' Props Collection Contents:
props("CodeLineNum") = True    ' Enable line numbering in code blocks
' Future extensions could include:
' props("ImageWidth") = 800     ' Max image width
' props("TableBorders") = True  ' Show table borders
' props("MathRenderer") = "LaTeX" ' Math formula format
```

### Props Flow Through System
1. **Creation**: `dView.props = New Collection` in main export functions
2. **Population**: `props.Add(CODE_LINE_NUM, "CodeLineNum")`
3. **Propagation**: Passed to `PrintTree(docTree)` → `PrintTree(node, props)`
4. **Usage**: Checked in PrintTree and passed to Code() function
5. **Application**: Controls line numbering in code block generation

### Props Usage Example in Code Processing
```basic
' In ViewHfm.Code() function:
If docView.props("CodeLineNum") Then
    Dim props As New Collection
    props.Add(True, "CodeLineNum")  ' Create local props for recursive call
    codeContent = docView.PrintTree(node, props)  ' Pass props down
End If
```

### Props Object Lifecycle
```
Main Function → Create Props → Populate → Pass to PrintTree → 
Recursive Calls → Code Processing → Line Number Control
```

## Code Block Line Numbering Control

### Current Implementation:
**Global Setting**: `CODE_LINE_NUM = True` in `DocModel.bas`

### To Disable Line Numbering:

**Method 1 - Global Disable:**
```basic
' In DocModel.bas, change:
Const CODE_LINE_NUM = False  ' Disable for all code blocks
```

**Method 2 - Runtime Control:**
```basic
' In MakeDocHfmView/MakeDocHtmlView functions:
dView.props = New Collection
With dView.props
    .Add(False, "CodeLineNum")  ' Disable line numbering
End With
```

**Method 3 - Per-Block Control:**
Modify `Code()` function in ViewHfm.bas:
```basic
Function Code(ByRef node)
    ' Force disable line numbering for specific styles
    If InStr(node.name_, "nolines") > 0 Then
        Code = docView.PrintTree(node)  ' No line numbers
    ElseIf docView.props("CodeLineNum") Then
        ' ... existing logic
    End If
End Function
```

### Line Numbering Flow:
1. **Check Global**: `CODE_LINE_NUM` constant sets default
2. **Check Props**: `props("CodeLineNum")` overrides in PrintTree
3. **Apply Format**: `Format_Num()` in Utils.bas formats line numbers
4. **Output**: Lines prefixed with formatted numbers (e.g., "   1 ", "   2 ")

### Props Integration with Line Numbering:
```basic
' In PrintTree function:
If Not IsMissing(props) Then
    If Not props Is Nothing Then
        On Error Resume Next
        If props("CodeLineNum") Then lineNum = 1  ' Enable numbering
        On Error GoTo 0
    End If
End If

' Later in paragraph processing:
If lineNum > 0 Then
    s = s & PrintNodePara(child, lineNum)  ' Pass line number
    lineNum = lineNum + 1                  ' Increment for next line
End If
```

## Style Recognition Patterns

### Code Blocks:
- **Pattern**: Paragraph style starting with `"code_"`
- **Examples**: `"code_cpp"`, `"code_python"`, `"code_javascript"`
- **Processing**: Extract language from style name, apply syntax highlighting markers

### Other Styles:
- **Quotations**: Style name = `"Quotations"` → blockquote formatting
- **Headings**: Style name starts with `"Heading"` → markdown headers
- **Default**: All other styles → standard paragraph formatting

## Memory Management

**Document Content Flow:**
```
LibreOffice Objects → Node Tree → Formatted Strings → Output File
```

**Key Variables:**
- `docTree`: Complete parsed document structure (Node tree)
- `viewAdapter`: Format-specific output handler (ViewHfm/ViewHtml classes)
- `props`: Runtime formatting configuration (Collection object)
- Node collections: Managed automatically by LibreOffice Basic garbage collector

**Props Object Memory Structure:**
```
Collection Object (props)
├── Key: "CodeLineNum" → Value: True/False
├── Key: "ImageWidth" → Value: Integer (future)
└── Key: "TableStyle" → Value: String (future)
```

**Document Order Preservation:**
The system maintains document element order through:
1. Sequential parsing via `paraEnum.nextElement()`
2. Ordered insertion into Collection objects (`children.Add(node)`)
3. Sequential iteration in PrintTree (`For Each child In node.children`)
4. Concatenated string output preserving original document flow

## Image Processing System

### Overview
The image processing system handles both remote URLs and embedded images, automatically extracting and copying embedded images to the `./img/` folder for proper markdown rendering.

### Image Processing Logic

#### 1. Image Type Detection
```basic
Function ProcessImage(ByRef imageObj, ByRef docURL As String) As String
    ' Check if image is remote URL (http/https) or embedded
    If Left(LCase(imageName), 4) = "http" Then
        ' Use remote URL directly
        ProcessImage = "![" & altText & "](" & imageName & ")"
    Else
        ' Process embedded image - extract and copy
        ' Apply filename cleaning and copy to ./img/ folder
        ProcessImage = "![" & altText & "](./img/" & fileName & ")"
    End If
End Function
```

#### 2. Image URL Sources
The system attempts to extract image URLs from multiple LibreOffice properties:
- `imageObj.Graphic.OriginURL` - Primary source for image location
- `imageObj.GraphicURL` - Alternative property for image URL
- Fallback handling for missing or inaccessible URLs

#### 3. Filename Cleaning Rules
Applied to all embedded images to ensure markdown compatibility:
- **Lowercase conversion**: All filenames converted to lowercase
- **Bracket removal**: `(` replaced with `-`, `)` removed entirely
- **Space replacement**: Spaces replaced with `-` characters
- **Path format**: Uses relative path `./img/filename.png`

**Example transformations**:
- `Image(1).PNG` → `image-1.png`
- `My Photo.jpg` → `my-photo.jpg`
- `diagram (final).gif` → `diagram-final.gif`

#### 4. File Copying Process
```basic
Function CopyImageFile(ByRef sourceURL As String, ByRef targetDir As String, ByRef fileName As String) As Boolean
    ' Convert LibreOffice URL to file system path
    Dim sourcePath As String : sourcePath = ConvertFromURL(sourceURL)
    
    ' Create ./img/ directory if it doesn't exist
    Dim imgDir As String : imgDir = targetDir & "img"
    If Not fso.FolderExists(imgDir) Then fso.CreateFolder(imgDir)
    
    ' Copy file with overwrite enabled
    fso.CopyFile sourcePath, targetPath, True
End Function
```

### Integration Points

#### 1. Header Image Processing
- **Location**: First page header content
- **Function**: `ProcessHeaderImage()` in DocModel.bas
- **Output**: Prepended to document with double line break
- **Alt text**: Uses "logo" as default if image title is empty

#### 2. Document Image Processing
- **Paragraph-anchored images**: Processed via `Image()` function
- **Inline images**: Processed via `InlineImage()` function
- **Integration**: Both functions use `ProcessImage()` for consistent handling

#### 3. Image Reference Generation
**Markdown format**:
```markdown
![alt-text](./img/filename.png)
```

**HTML format** (for inline images):
```html
<img inline="true" src="./img/filename.png" />
```

### Directory Structure
```
document-folder/
├── document.odt
├── document_hfm.md
└── img/
    ├── header-image.png
    ├── diagram-1.jpg
    └── photo-final.gif
```

### Error Handling
- **Missing images**: Fallback to `./img/missing-image.png`
- **URL extraction failure**: Uses generic filename based on image type
- **File copy errors**: Continues processing with reference to intended location
- **Directory creation**: Automatically creates `./img/` folder if needed

### Remote URL Handling
- **HTTP/HTTPS URLs**: Used directly without modification
- **No local copying**: Remote images remain as external references
- **Bandwidth consideration**: Remote images require internet access for viewing

### Image Processing Flow
```
LibreOffice Image Object
├── Extract URL/Properties
├── Determine Type (Remote vs Embedded)
├── Remote URL → Use directly
└── Embedded Image
    ├── Clean filename
    ├── Copy to ./img/ folder
    └── Generate relative path reference
```

### Future Enhancements
- **Image optimization**: Resize large images for web use
- **Format conversion**: Convert unsupported formats to web-friendly formats
- **Batch processing**: Handle multiple images efficiently
- **Error reporting**: Detailed logging of image processing issues
- **Alternative text generation**: Auto-generate alt text from image content analysis
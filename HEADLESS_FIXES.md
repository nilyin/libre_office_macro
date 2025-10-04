# Headless Mode Conversion Issues - Root Cause Analysis and Fixes

## Root Cause Analysis

### Primary Issues Identified

1. **Path Separator Detection Failure in Headless Mode**
   - `GetPathSeparator()` function in `Utils.bas` relies on `Environ("WINDIR")` which may not work in headless Linux
   - Results in wrong path separators being used (`\` instead of `/`)
   - Causes file operations to fail silently

2. **Image Folder Name Generation Failure**
   - `GenerateImageFolderName()` function fails to extract filename properly
   - Results in folder names like `img_` instead of `img_filename`
   - Indicates `ConvertFromURL()` or path parsing issues in headless mode

3. **File System Service Unavailability**
   - `CreateObject("Scripting.FileSystemObject")` may not be available in headless Linux
   - Directory creation and file operations fail silently
   - No error reporting in current implementation

4. **LibreOffice Document Loading Issues**
   - `StarDesktop.loadComponentFromURL()` may behave differently in headless mode
   - Document properties and URL handling may be inconsistent
   - Hidden document loading may not initialize all services

## Specific Code Issues

### 1. Utils.bas - GetPathSeparator() Function
```basic
Function GetPathSeparator() As String
    ' Current implementation fails in headless Linux
    On Error Resume Next
    Dim testPath As String : testPath = Environ("WINDIR")
    If Err.Number = 0 And testPath <> "" Then
        GetPathSeparator = "\"  ' Windows
    Else
        GetPathSeparator = "/"  ' Unix/Linux/macOS
    End If
    On Error GoTo 0
End Function
```

**Problem**: `Environ("WINDIR")` may not be reliable in headless mode

### 2. DocModel.bas - ExportDir Function
```basic
Dim fname As String : fname = Dir$(Folder + GetPathSeparator() + "*.odt", 0)
```

**Problem**: Path construction may use wrong separator, `Dir$()` behavior differs in headless mode

### 3. DocView.bas - Image Processing
```basic
Dim fso : fso = CreateObject("Scripting.FileSystemObject")
```

**Problem**: FSO may not be available in headless Linux environment

## Recommended Fixes

### Fix 1: Robust Path Separator Detection
```basic
Function GetPathSeparator() As String
    ' Multi-method approach for reliable detection
    On Error Resume Next
    
    ' Method 1: Check OS environment variables
    Dim winDir As String : winDir = Environ("WINDIR")
    If Err.Number = 0 And winDir <> "" Then
        GetPathSeparator = "\"
        Exit Function
    End If
    
    ' Method 2: Check Unix/Linux environment
    Dim homePath As String : homePath = Environ("HOME")
    If Err.Number = 0 And homePath <> "" Then
        GetPathSeparator = "/"
        Exit Function
    End If
    
    ' Method 3: Test path creation
    Dim testPath As String : testPath = "/tmp"
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    If Not IsEmpty(fso) And fso.FolderExists(testPath) Then
        GetPathSeparator = "/"
        Exit Function
    End If
    
    ' Method 4: Default based on LibreOffice platform detection
    Dim platform As String : platform = GetOS()
    If InStr(LCase(platform), "win") > 0 Then
        GetPathSeparator = "\"
    Else
        GetPathSeparator = "/"
    End If
    
    On Error GoTo 0
End Function

Function GetOS() As String
    ' Use LibreOffice's built-in OS detection
    Dim oConfigProvider : oConfigProvider = CreateUnoService("com.sun.star.configuration.ConfigurationProvider")
    Dim oPathSettings : oPathSettings = CreateUnoService("com.sun.star.util.PathSettings")
    GetOS = oPathSettings.Module
End Function
```

### Fix 2: Enhanced File System Operations
```basic
Function CreateDirectorySafe(ByRef dirPath As String) As Boolean
    On Error Resume Next
    CreateDirectorySafe = False
    
    ' Method 1: Try FileSystemObject (Windows/Wine)
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    If Not IsEmpty(fso) Then
        If Not fso.FolderExists(dirPath) Then
            fso.CreateFolder(dirPath)
        End If
        CreateDirectorySafe = fso.FolderExists(dirPath)
        If CreateDirectorySafe Then Exit Function
    End If
    
    ' Method 2: Try LibreOffice SimpleFileAccess (Cross-platform)
    Dim fileAccess : fileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not IsEmpty(fileAccess) Then
        Dim urlPath As String : urlPath = ConvertToURL(dirPath)
        If Not fileAccess.exists(urlPath) Then
            fileAccess.createFolder(urlPath)
        End If
        CreateDirectorySafe = fileAccess.exists(urlPath)
        If CreateDirectorySafe Then Exit Function
    End If
    
    ' Method 3: Try Shell command (Linux/Unix)
    If GetPathSeparator() = "/" Then
        Shell("mkdir -p """ & dirPath & """", 0)
        ' Check if directory was created
        Dim testFile As String : testFile = dirPath & "/.test"
        On Error Resume Next
        Open testFile For Output As #1
        Close #1
        Kill testFile
        CreateDirectorySafe = (Err.Number = 0)
    End If
    
    On Error GoTo 0
End Function
```

### Fix 3: Robust File Enumeration
```basic
Function GetODTFiles(ByRef folderPath As String) As Variant
    Dim fileList() As String
    Dim fileCount As Long : fileCount = 0
    
    On Error Resume Next
    
    ' Method 1: Try Dir$ function
    Dim fileName As String : fileName = Dir$(folderPath & GetPathSeparator() & "*.odt", 0)
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
                    If Right(LCase(fileUrl), 4) = ".odt" Then
                        ReDim Preserve fileList(fileCount)
                        fileList(fileCount) = ConvertFromURL(fileUrl)
                        fileCount = fileCount + 1
                    End If
                Next
            End If
        End If
    End If
    
    On Error GoTo 0
    
    If fileCount > 0 Then
        GetODTFiles = fileList
    Else
        GetODTFiles = Array() ' Empty array
    End If
End Function
```

### Fix 4: Enhanced ExportDir Function
```basic
Sub ExportDir(Folder As String, Optional Hfm As Variant)
    Dim useHfm As Boolean : useHfm = True
    If Not IsMissing(Hfm) Then
        If IsNumeric(Hfm) Then useHfm = CBool(Hfm)
        If VarType(Hfm) = vbBoolean Then useHfm = Hfm
    End If
    
    ' Enhanced error logging
    Dim logFile As String : logFile = Folder & GetPathSeparator() & "conversion.log"
    Dim logNum As Integer : logNum = FreeFile
    
    On Error Resume Next
    Open logFile For Output As #logNum
    Print #logNum, "=== LibreOffice Conversion Log ==="
    Print #logNum, "Timestamp: " & Now()
    Print #logNum, "Folder: " & Folder
    Print #logNum, "HFM Mode: " & useHfm
    Print #logNum, "Path Separator: " & GetPathSeparator()
    
    ' Get ODT files using robust method
    Dim odtFiles : odtFiles = GetODTFiles(Folder)
    Print #logNum, "ODT Files Found: " & UBound(odtFiles) + 1
    
    If UBound(odtFiles) >= 0 Then
        Dim Props(0) As New com.sun.star.beans.PropertyValue
        Props(0).Name = "Hidden"
        Props(0).Value = True
        
        Dim Comp As Object
        Dim i As Long
        For i = 0 To UBound(odtFiles)
            Dim filePath As String : filePath = odtFiles(i)
            Print #logNum, "Processing: " & filePath
            
            Dim url As String : url = ConvertToURL(filePath)
            Print #logNum, "URL: " & url
            
            Comp = StarDesktop.loadComponentFromURL(url, "_blank", 0, Props)
            If Not IsEmpty(Comp) Then
                Print #logNum, "Document loaded successfully"
                
                If useHfm Then
                    MakeDocHfmView Comp
                    Print #logNum, "HFM conversion completed"
                Else
                    MakeDocHtmlView Comp
                    Print #logNum, "HTML conversion completed"
                End If
                
                Comp.close(True)
                Print #logNum, "Document closed"
            Else
                Print #logNum, "ERROR: Failed to load document"
            End If
        Next
    Else
        Print #logNum, "ERROR: No ODT files found"
    End If
    
    Print #logNum, "=== Conversion Complete ==="
    Close #logNum
    On Error GoTo 0
End Sub
```

### Fix 5: Enhanced Image Processing
```basic
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
    If InStr(fileName, ".") > 0 Then
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    End If
    
    ' Fallback to generic name if all methods fail
    If fileName = "" Then fileName = "document"
    
    GenerateImageFolderName = "img_" & fileName
    On Error GoTo 0
End Function
```

## Implementation Priority

1. **CRITICAL**: Fix `GetPathSeparator()` function for proper path handling
2. **HIGH**: Implement robust file system operations with fallbacks
3. **HIGH**: Fix `GenerateImageFolderName()` with multiple extraction methods
4. **MEDIUM**: Add comprehensive error logging for debugging
5. **MEDIUM**: Enhance `ExportDir` function with better error handling
6. **LOW**: Optimize performance for batch processing

## Testing Strategy

### 1. Headless Mode Testing
```bash
# Test individual file conversion
soffice --invisible --nofirststartwizard --headless --norestore \
  "macro:///DocExport.DocModel.MakeDocHfmView" \
  --invisible file:///path/to/test.odt

# Test batch conversion
soffice --invisible --nofirststartwizard --headless --norestore \
  "macro:///DocExport.DocModel.ExportDir(\"/path/to/odt/files\",1)"
```

### 2. Debug Information Collection
- Check conversion.log files for detailed error information
- Verify path separator detection
- Confirm file enumeration results
- Validate image folder creation

### 3. Cross-Platform Validation
- Test on Linux headless environment
- Test on Windows headless environment
- Verify path handling consistency
- Confirm image extraction works

## Expected Results After Fixes

1. **Successful Batch Conversion**: All ODT files should convert without errors
2. **Proper Image Folders**: Folders named `img_filename` should be created correctly
3. **Cross-Platform Compatibility**: Same behavior on Windows and Linux
4. **Detailed Logging**: Clear error messages for troubleshooting
5. **Robust Error Handling**: Graceful fallbacks when services are unavailable

## Deployment Steps

1. Apply fixes to all affected `.bas` files
2. Update the macro package using `update_macros.py`
3. Test in GUI mode first to ensure no regressions
4. Test in headless mode with single file
5. Test batch conversion with multiple files
6. Validate image processing and folder creation
7. Deploy to production environment
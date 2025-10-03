# Cross-Platform Path Separator Fixes

## Problem Description
When testing on Linux, the image folder naming was incorrect:
- **Expected**: `img_RU.ECO.00101-01_90_001-0/`
- **Actual**: `img_/tmp/tmp.9Wgb8RyaaC/RU.ECO.00101-01_90_001-0/`

The issue occurred in both individual file export and bulk export (`ExportDir`) operations.

## Root Causes

### 1. Hardcoded Path Separators
The `GenerateImageFolderName()` function in `DocModel.bas` was using hardcoded backslash `\` instead of the platform-specific path separator.

**Problem Code**:
```basic
Dim fileName As String : fileName = Mid(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), "\") + 1)
```

**Fixed Code**:
```basic
Dim fileName As String : fileName = Mid(ConvertFromURL(docURL), InStrRev(ConvertFromURL(docURL), GetPathSeparator()) + 1)
```

### 2. Wrong Document URL in Batch Processing
In `ViewHfm.bas` and `ViewHtml.bas`, the `Image()` and `InlineImage()` functions were using `ThisComponent.URL` which refers to the currently active document in the UI, not the document being processed in batch mode.

**Problem Code**:
```basic
Function Image(ByRef lo)
    Dim imageUrl As String : imageUrl = docView.ProcessImage(lo, ThisComponent.URL)
    Image = imageUrl & "  " & CHR$(10)
End Function
```

**Fixed Code**:
```basic
Function Image(ByRef lo)
    Dim imageUrl As String : imageUrl = docView.ProcessImage(lo, docView.docURL)
    Image = imageUrl & "  " & CHR$(10)
End Function
```

## Changes Made

### 1. DocModel.bas
- Fixed `GenerateImageFolderName()` to use `GetPathSeparator()` instead of hardcoded `\`
- Added `dView.docURL = doc.URL` in both `MakeDocHtmlView()` and `MakeDocHfmView()` functions

### 2. DocView.bas
- Added new public property: `Public docURL As String`
- This stores the correct document URL for image processing

### 3. ViewHfm.bas
- Updated `Image()` function to use `docView.docURL` instead of `ThisComponent.URL`
- Updated `InlineImage()` function to use `docView.docURL` instead of `ThisComponent.URL`

### 4. ViewHtml.bas
- Updated `InlineImage()` function to use `docView.docURL` instead of `ThisComponent.URL`

### 5. Utils.bas
- Already contained `GetPathSeparator()` function that detects platform:
  - Returns `\` for Windows
  - Returns `/` for Unix/Linux/macOS

## Testing Scenarios

### Individual File Export
- **Windows**: `C:\Documents\report.odt` → `img_report/`
- **Linux**: `/home/user/report.odt` → `img_report/`
- **macOS**: `/Users/user/report.odt` → `img_report/`

### Batch Export (ExportDir)
- **Windows**: `macro:///DocExport.DocModel.ExportDir("D:\odt",1)`
- **Linux**: `macro:///DocExport.DocModel.ExportDir("/tmp/odt",1)`
- **macOS**: `macro:///DocExport.DocModel.ExportDir("/Users/user/odt",1)`

All scenarios now correctly extract the filename and create properly named image folders.

## Benefits
- **Cross-platform compatibility**: Works correctly on Windows, Linux, and macOS
- **Batch processing**: Correctly handles multiple documents in `ExportDir`
- **Consistent naming**: Image folders always use pattern `img_{filename}`
- **No path pollution**: Folder names no longer include directory paths

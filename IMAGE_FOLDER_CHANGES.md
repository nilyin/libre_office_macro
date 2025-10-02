# Image Folder Naming Changes

## Summary
Modified the LibreOffice macro to change image folder naming from fixed 'img' to dynamic 'img_' + source ODT filename pattern.

## Changes Made

### 1. DocModel.bas
- Added `GenerateImageFolderName()` function to create dynamic folder names
- Modified `ProcessHeaderImage()` function to use dynamic folder naming
- Updated folder creation and markdown reference generation

### 2. DocView.bas  
- Added `GenerateImageFolderName()` function (duplicate for modularity)
- Modified `ExtractImageFile()` function signature to accept docURL parameter
- Modified `CopyImageFile()` function signature to accept docURL parameter
- Updated `ProcessImage()` function to use dynamic folder names in both extraction and markdown generation

## Folder Naming Pattern
- **Before**: Fixed folder name `img`
- **After**: Dynamic folder name `img_` + `{source_odt_filename_without_extension}`

## Examples
- Source file: `my_document.odt` → Image folder: `img_my_document`
- Source file: `article-draft.odt` → Image folder: `img_article-draft`
- Source file: `report_2024.odt` → Image folder: `img_report_2024`

## Impact
- Each ODT file now generates its own uniquely named image folder
- Prevents image conflicts when multiple ODT files are processed in the same directory
- Maintains backward compatibility with existing markdown references
- Works for both HFM (Habr Flavored Markdown) and HTML export formats

## Functions Modified
1. `DocModel.ProcessHeaderImage()` - Header image processing
2. `DocView.ExtractImageFile()` - Embedded image extraction  
3. `DocView.CopyImageFile()` - External image copying
4. `DocView.ProcessImage()` - Main image processing and markdown generation

All changes maintain the existing functionality while implementing the new dynamic folder naming requirement.
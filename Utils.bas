REM Author: Dmitry A. Borisov, ddaabb@mail.ru (CC BY 4.0)
Option Explicit

' Format line number with fixed width padding for code blocks
' @param lineNum: Line number to format
' @param Size: Optional width for padding (default: 4)
' @return: Formatted line number string with right padding
Function Format_Num(ByVal lineNum As Integer, Optional Size As Variant) As String
    Dim sz As Integer : sz = 4  ' Default width
    If Not IsMissing(Size) Then
        If IsNumeric(Size) Then sz = CInt(Size)
    End If
    Dim s As String : s = "" + lineNum  ' Convert number to string
    s = s + String(sz - Len(s), " ")     ' Add right padding with spaces
    Format_Num = s
End Function

' Escape HTML special characters in text
' @param s: Input string to escape
' @return: String with HTML entities for < and > characters
Function Escape_Characters(ByRef s As String) As String
    Dim sz As Integer, c As String,  i As Integer, r As String : r = ""  ' Variables for processing
    sz = Len(s)  ' Get string length
    i = 1        ' Start at position 1 (VBA/Basic uses 1-based indexing)
    
    ' Process each character in the string
    While i <= sz
        c = Mid(s, i, 1)  ' Get current character
        Select Case c
            Case "<"
                r = r + "&lt;"  ' Replace < with HTML entity
            Case ">"
                r = r + "&gt;"  ' Replace > with HTML entity
            Case Else
                r = r + c       ' Keep other characters as-is
        End Select
        i = i + 1  ' Move to next character
    Wend
    Escape_Characters = r
End Function

' Debug logging function for headless mode
Sub LogDebug(ByRef message As String, Optional ByRef logFile As String)
    On Error Resume Next
    If IsMissing(logFile) Or logFile = "" Then logFile = "/tmp/libreoffice_debug.log"
    
    Dim fileNum As Integer : fileNum = FreeFile
    Open logFile For Append As #fileNum
    Print #fileNum, Now() & " - " & message
    Close #fileNum
    On Error GoTo 0
End Sub

' Test subroutine for utility functions (for development/testing purposes)
Sub Main
    Dim s As String : s = "'" + Format_Num(43, 12) + "'"  ' Test Format_Num function
    s = "<?xml version=""1.0"" encoding=""UTF-8""?>"      ' Test XML string
    s = Escape_Characters(s)  ' Test character escaping
End Sub

' Get platform-specific path separator with robust detection
' @return: "\" for Windows, "/" for Unix/Linux/macOS
Function GetPathSeparator() As String
    On Error Resume Next
    
    ' Method 1: Check Windows environment
    Dim winDir As String : winDir = Environ("WINDIR")
    If Err.Number = 0 And winDir <> "" Then
        GetPathSeparator = "\"
        Exit Function
    End If
    
    ' Method 2: Check Unix/Linux environment
    Dim homePath As String : homePath = Environ("HOME")
    If Err.Number = 0 And homePath <> "" And Left(homePath, 1) = "/" Then
        GetPathSeparator = "/"
        Exit Function
    End If
    
    ' Method 3: Test LibreOffice platform detection
    Dim oPathSettings : oPathSettings = CreateUnoService("com.sun.star.util.PathSettings")
    If Not IsEmpty(oPathSettings) Then
        Dim tempPath As String : tempPath = oPathSettings.Temp
        If tempPath <> "" Then
            If InStr(tempPath, "/") > 0 Then
                GetPathSeparator = "/"
                Exit Function
            ElseIf InStr(tempPath, "\") > 0 Then
                GetPathSeparator = "\"
                Exit Function
            End If
        End If
    End If
    
    ' Default to Unix separator for headless environments
    GetPathSeparator = "/"
    On Error GoTo 0
End Function

' Enhanced file system operations with cross-platform support
Function CreateDirectorySafe(ByRef dirPath As String) As Boolean
    On Error Resume Next
    CreateDirectorySafe = False
    
    ' Method 1: Try LibreOffice SimpleFileAccess (Cross-platform)
    Dim fileAccess : fileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not IsEmpty(fileAccess) Then
        Dim urlPath As String : urlPath = ConvertToURL(dirPath)
        If Not fileAccess.exists(urlPath) Then
            fileAccess.createFolder(urlPath)
        End If
        CreateDirectorySafe = fileAccess.exists(urlPath)
        If CreateDirectorySafe Then Exit Function
    End If
    
    ' Method 2: Try FileSystemObject (Windows/Wine)
    Dim fso : fso = CreateObject("Scripting.FileSystemObject")
    If Not IsEmpty(fso) Then
        If Not fso.FolderExists(dirPath) Then
            fso.CreateFolder(dirPath)
        End If
        CreateDirectorySafe = fso.FolderExists(dirPath)
    End If
    
    On Error GoTo 0
End Function
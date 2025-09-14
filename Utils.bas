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

' Test subroutine for utility functions (for development/testing purposes)
Sub Main
    Dim s As String : s = "'" + Format_Num(43, 12) + "'"  ' Test Format_Num function
    s = "<?xml version=""1.0"" encoding=""UTF-8""?>"      ' Test XML string
    s = Escape_Characters(s)  ' Test character escaping
End Sub

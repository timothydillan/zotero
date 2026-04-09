Attribute VB_Name = "ZoteroBatchUpdate"
' ZoteroBatchUpdate - Batch field update module for Zotero
' This module reads a batch file of field updates and applies them
' all in-process, avoiding per-field IPC overhead with Zotero.
'
' Batch file format:
'   Line 1: total number of update blocks
'   For each block:
'     Header line: FIELD_INDEX<TAB>ACTION
'       ACTION: T = set text, C = set code, D = delete field, R = remove code
'     Data line: the data (single line, with \n escaped as \\n and \ escaped as \\)
'
' Field indices are 0-based Zotero indices into the filtered list of
' Zotero fields (those whose code contains "ZOTERO").

Option Explicit

Public Sub ZoteroBatchUpdateFromFile(filePath As String)
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 1. Collect all Zotero fields in document order
    Dim doc As Document
    Set doc = ActiveDocument

    Dim zoteroFields As New Collection
    Dim f As Field
    For Each f In doc.Fields
        If InStr(f.Code.Text, "ZOTERO") > 0 Then
            zoteroFields.Add f
        End If
    Next f

    If zoteroFields.Count = 0 Then
        GoTo Cleanup
    End If

    ' 2. Read the batch file
    Dim fNum As Integer
    fNum = FreeFile
    Open filePath For Input As #fNum

    ' Read number of updates
    Dim numUpdates As Long
    Dim headerLine As String
    Line Input #fNum, headerLine
    numUpdates = CLng(headerLine)

    ' 3. Process each update
    Dim i As Long
    Dim parts() As String
    Dim fieldIdx As Long
    Dim action As String
    Dim dataLine As String
    Dim unescapedData As String

    For i = 1 To numUpdates
        If EOF(fNum) Then Exit For

        ' Read header: FIELD_INDEX<TAB>ACTION
        Line Input #fNum, headerLine
        parts = Split(headerLine, vbTab)

        If UBound(parts) < 1 Then GoTo NextUpdate

        fieldIdx = CLng(parts(0))
        action = parts(1)

        ' Read data line
        If EOF(fNum) Then Exit For
        Line Input #fNum, dataLine

        ' Unescape: \\n -> newline, \\\\ -> backslash
        unescapedData = UnescapeData(dataLine)

        ' Convert 0-based Zotero index to 1-based VBA Collection index
        Dim vbaIdx As Long
        vbaIdx = fieldIdx + 1

        If vbaIdx < 1 Or vbaIdx > zoteroFields.Count Then GoTo NextUpdate

        ' Apply the update
        Select Case action
            Case "T"
                ' Set field result text
                Dim targetField As Field
                Set targetField = zoteroFields(vbaIdx)

                ' Check if this is RTF content
                If Left(unescapedData, 5) = "{\rtf" Then
                    ' Write RTF to temp file and insert
                    Dim rtfPath As String
                    rtfPath = Environ("TMPDIR") & "zotero_rtf_temp.rtf"
                    Dim rtfNum As Integer
                    rtfNum = FreeFile
                    Open rtfPath For Output As #rtfNum
                    Print #rtfNum, unescapedData
                    Close #rtfNum

                    targetField.Result.Select
                    Selection.InsertFile rtfPath

                    ' Clean up temp file
                    Kill rtfPath
                Else
                    ' Plain text - direct assignment
                    targetField.Result.Text = unescapedData
                End If

            Case "C"
                ' Set field code (always plain text)
                ' Field code needs leading/trailing spaces in Word
                zoteroFields(vbaIdx).Code.Text = " " & unescapedData & " "

            Case "D"
                ' Delete the field
                zoteroFields(vbaIdx).Delete

            Case "R"
                ' Remove field code (unlink field, keep text)
                zoteroFields(vbaIdx).Unlink
        End Select

NextUpdate:
    Next i

    Close #fNum

Cleanup:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Write completion signal
    Dim sigPath As String
    sigPath = filePath & ".done"
    Dim sigNum As Integer
    sigNum = FreeFile
    Open sigPath For Output As #sigNum
    Print #sigNum, "OK"
    Close #sigNum

    Exit Sub

ErrorHandler:
    ' Write error signal
    Dim errPath As String
    errPath = filePath & ".error"
    Dim errNum As Integer
    errNum = FreeFile
    Open errPath For Output As #errNum
    Print #errNum, "Error " & Err.Number & ": " & Err.Description
    Close #errNum

    Resume Cleanup
End Sub

Private Function UnescapeData(s As String) As String
    ' Unescape \\n -> newline, \\\\ -> \
    ' Process in order: \\\\ first, then \\n
    Dim result As String
    result = s

    ' Use a placeholder for literal backslash to avoid double-processing
    result = Replace(result, "\\\\", Chr(0))
    result = Replace(result, "\\n", vbLf)
    result = Replace(result, Chr(0), "\")

    UnescapeData = result
End Function

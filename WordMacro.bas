Sub ExportDocumentStructure()
'
' Get the document structure (for example, all the headings) and export it to a file, so that it can be merged with the PDF created from macOS.
'
'
' par will be the paragraph that is currently being read
Dim par As Word.Paragraph
' correctRange will be the subrange of the paragraph that excludes page breaks
Dim correctRange As Word.Range
' skipTheseCharacters will number, that is set to 0 at every paragraph iteration, that tells the script when the first valid character (so, not page breaks) actually is
Dim skipTheseCharacters As Long
' str will be our output string
Dim str As String

            
' Add title and author on the top of the file so that we can add also these two metadata.
str = Word.ActiveDocument.BuiltInDocumentProperties("Title") & vbNewLine & Word.ActiveDocument.BuiltInDocumentProperties("Author") & vbNewLine
For Each par In Word.ActiveDocument.Paragraphs
    If par.OutlineLevel < 10 Then
        ' We'll now count the number of characters at the start of the paragraph text that are either page breaks, column breaks or line breaks so that we can skip them in the paragraph position calculation
        skipTheseCharacters = 0
        Do While skipTheseCharacters < Len(par.Range.text)
            Select Case Mid$(par.Range.text, skipTheseCharacters + 1, 1)
                Case Chr(12), Chr(14), Chr(11)
                    skipTheseCharacters = skipTheseCharacters + 1
                Case Else
                    Exit Do
            End Select
        Loop
        Set correctRange = par.Range
        correctRange.SetRange par.Range.Start + skipTheseCharacters, par.Range.Start + skipTheseCharacters   ' collapsed at first real character
        str = str & CStr(par.OutlineLevel) & ";" & CStr(correctRange.Information(wdActiveEndPageNumber)) & ";" & CStr(correctRange.Information(wdHorizontalPositionRelativeToPage)) & ";" & CStr(correctRange.Information(wdVerticalPositionRelativeToPage)) & ";" & Replace(Replace(par.Range.text, Chr(12), ""), Chr(14), "") & vbNewLine
    End If
Next par


' Now let's save the file. Tbf the following code was made by pasting different parts found online

Dim fileNum As Integer
fileNum = FreeFile
Dim savePath As String
' Use MacScript to open the Save File Picker on macOS
savePath = MacScript("return (choose file name with prompt " & _
        Chr(34) & "Save text file as:" & Chr(34) & _
        " default name " & Chr(34) & ActiveDocument.Name & "-Titles.txt" & Chr(34) & _
        ") as string")
savePath = MacScript("return POSIX path of (" & Chr(34) & savePath & Chr(34) & ")")

Open (savePath) For Output As #fileNum
    Print #fileNum, str
    Close #fileNum

End Sub

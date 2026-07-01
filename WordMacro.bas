Sub ExportDocumentStructure()
'
' Get the document structure (for example, all the headings) and export it to a file, so that it can be merged with the PDF created from macOS.
'
'

' str will be our output string
Dim str As String
' Add title and author on the top of the file so that we can also add these two metadata.
str = Word.ActiveDocument.BuiltInDocumentProperties("Title") & vbNewLine & Word.ActiveDocument.BuiltInDocumentProperties("Author") & vbNewLine


For Each par In Word.ActiveDocument.Paragraphs
    If par.OutlineLevel < 10 Then ' OutlineLevel = 10 is body text
    ' Separator char: ;
    ' Syntax: Outline level; Page; Horizontal position; Vertical position; Text
        str = str & CStr(par.OutlineLevel) & ";" & CStr(par.Range.Information(wdActiveEndPageNumber)) & ";" & CStr(par.Range.Information(wdHorizontalPositionRelativeToPage)) & ";" & CStr(par.Range.Information(wdVerticalPositionRelativeToPage)) & ";" & par.Range.Text & vbNewLine ' It seems that par.Range.Text automatically removes new lines without space, so we don't need to sanitize them with \n
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


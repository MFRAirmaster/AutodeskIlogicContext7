' Title: Have the sheet name automatically generate in Title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336#messageview_0
' Category: api
' Scraped: 2025-10-07T12:53:24.187979

Dim doc As DrawingDocument = ThisDoc.Document
Dim sheet As Sheet = doc.ActiveSheet
Dim titleBlock As TitleBlock = sheet.TitleBlock

Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
        Cast(Of Inventor.TextBox).
        Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
        Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
        FirstOrDefault()
If (textBox Is Nothing) Then Throw New Exception("Promted entry not found.")
titleBlock.SetPromptResultText(textBox, sheet.Name)
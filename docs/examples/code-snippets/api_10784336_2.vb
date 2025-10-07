' Title: Have the sheet name automatically generate in Title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336#messageview_0
' Category: api
' Scraped: 2025-10-07T12:53:24.187979

Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes.
	    Cast(Of Inventor.TextBox).
	    Where(Function(tb) tb.FormattedText.Contains("<Prompt")).
	    Where(Function(tb) tb.FormattedText.Contains("SHEET DESCRIPTION")).
	    FirstOrDefault()
    If (textBox Is Nothing) Then 
		MsgBox("Promted entry not found on sheet: " & Sheet.Name)
		continue for
	End If

    Dim sheetDescription = Sheet.Name
    Dim puntPlace = InStr(sheetDescription, ":") - 1
    sheetDescription = sheetDescription.Substring(0, puntPlace)
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
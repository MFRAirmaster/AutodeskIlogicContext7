' Title: Have the sheet name automatically generate in Title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/have-the-sheet-name-automatically-generate-in-title-block/td-p/10784336#messageview_0
' Category: api
' Scraped: 2025-10-09T09:08:43.934542

Dim doc As DrawingDocument = ThisDoc.Document
For Each sheet As Sheet In doc.Sheets
    Dim titleBlock As TitleBlock = Sheet.TitleBlock

    ' Find the textbox within the title block's sketch that contains the prompt markers
    Dim textBox As Inventor.TextBox = titleBlock.Definition.Sketch.TextBoxes _
        .Cast(Of Inventor.TextBox)() _
        .Where(Function(tb) tb.FormattedText.Contains("<Prompt") AndAlso tb.FormattedText.Contains("SHEET TITLE")) _
        .FirstOrDefault()

    If (textBox Is Nothing) Then
        MsgBox("Prompted entry not found on sheet: " & Sheet.Name)
        Continue For
    End If

    ' Extract the sheet description; here, trimming after colon
    Dim sheetDescription As String = Sheet.Name
    Dim puntPlace As Integer = InStr(sheetDescription, ":")
    If puntPlace > 0 Then
        sheetDescription = sheetDescription.Substring(0, puntPlace - 1)
    End If

    ' Set the prompt result text in the title block
    titleBlock.SetPromptResultText(textBox, sheetDescription)
Next
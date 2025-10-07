' Title: Add leading zeros to title block
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/add-leading-zeros-to-title-block/td-p/9722111#messageview_0
' Category: api
' Scraped: 2025-10-07T14:09:09.528504

Dim oSheets As Sheets = ThisDrawing.Document.Sheets
For Each oSheet As Sheet In oSheets
	oSheetNumber = CInt(oSheet.Name.Split(":")(1)).ToString("00")
	Dim oTB As TitleBlock = oSheet.TitleBlock
	For Each oTextBox As Inventor.TextBox In oTB.Definition.Sketch.TextBoxes
		If oTextBox.Text = "<CustomSheetNumber>"
			oTB.SetPromptResultText(oTextBox, (oSheets.Count.ToString("00") & oSheetNumber))
			Exit For
		End If
	Next
Next
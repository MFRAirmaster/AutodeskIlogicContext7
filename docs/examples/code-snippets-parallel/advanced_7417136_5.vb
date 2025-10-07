' Title: update drawing line colour based on part colour
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-line-colour-based-on-part-colour/td-p/7417136
' Category: advanced
' Scraped: 2025-10-07T13:59:00.967123

Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisApplication.ActiveDocument
Dim oSheets As Sheets
oSheets = ThisDoc.Document.Sheets

Dim oSheet As Sheet
'Iterate through the sheets
i=0
For Each oSheet In oSheets		
	If oSheet.Name.Contains("Stückliste") Then
		i = i + 1
	End If	
Next

oPartsListSheetCount = i

i = 0
For Each oSheet In oSheets	
	If oSheet.Name.Contains("Stückliste") Then
		If oSheet.TitleBlock IsNot Nothing Then 
		    oTitleBlock = oSheet.TitleBlock
		    oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
		    For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
				If oTextBox.Text = "BlattAnzahl"
					oPromptEntry1 = oPartsListSheetCount
					Call oTitleBlock.SetPromptResultText(oTextBox, oPromptEntry1)
					
				End If
				If oTextBox.Text = "BlattNummer"
					i = i + 1
					oPromptEntry2 =  i
					Call oTitleBlock.SetPromptResultText(oTextBox, oPromptEntry2)
					
				End If
		    Next
		End If
	End If
Next
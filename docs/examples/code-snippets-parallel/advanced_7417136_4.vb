' Title: update drawing line colour based on part colour
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-line-colour-based-on-part-colour/td-p/7417136
' Category: advanced
' Scraped: 2025-10-09T09:03:39.887945

Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisApplication.ActiveDocument
Dim oSheets As Sheets
oSheets = ThisDoc.Document.Sheets

Dim oSheet As Sheet
'Iterate through the sheets
i=0
For Each oSheet In oSheets		
	If oSheet.ExcludeFromCount = True Then
		i = i + 1
		oSheet.Name = "PartsList" & "-" & i		
	End If	
Next

oPartsListSheetCount = i

i = 0
For Each oSheet In oSheets	
	If oSheet.Name.Contains("PartsList") Then
		If oSheet.TitleBlock IsNot Nothing Then 
		    oTitleBlock = oSheet.TitleBlock
		    oTextBoxes = oTitleBlock.Definition.Sketch.TextBoxes
		    For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
				If oTextBox.Text = "PartsList_OF_PartsLists"
					i = i + 1
					oPromptEntry =  i & " OF " & oPartsListSheetCount
					Call oTitleBlock.SetPromptResultText(oTextBox, oPromptEntry)
					
				End If
		    Next
		End If
	End If
Next
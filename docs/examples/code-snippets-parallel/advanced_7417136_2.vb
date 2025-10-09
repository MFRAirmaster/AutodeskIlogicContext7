' Title: update drawing line colour based on part colour
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-line-colour-based-on-part-colour/td-p/7417136
' Category: advanced
' Scraped: 2025-10-09T09:03:39.887945

Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisApplication.ActiveDocument
Dim oSheets As Sheets
oSheets = ThisDoc.Document.Sheets
Dim oSheet As Sheet
oSheet = ThisDoc.Document.ActiveSheet
Dim oPageTotal As Integer
i = oPageTotal

'Iterate through the sheets
For Each oSheet In oSheets
	oSheet.Activate
		
	'Verify that the Sheet Name is Page
	If oSheet.Name = "Page" Then
		i = i+1
	End If
Next

iProperties.Value("Custom", "PartslistSheetTotal") = oPageTotal
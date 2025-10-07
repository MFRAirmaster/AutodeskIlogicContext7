' Title: Excluding Print and Count if Sheet name matches
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/excluding-print-and-count-if-sheet-name-matches/td-p/11396697
' Category: api
' Scraped: 2025-10-07T13:28:29.992552

Dim oDDoc As DrawingDocument = ThisDrawing.Document
Dim oExcludeSheets As Boolean = InputRadioBox("Client or All Sheets?", "Client", "All", Client, "Toggle Sheets")
For Each oSheet As Sheet In oDDoc.Sheets
	If Not oSheet.Name.StartsWith("C") Then
		oSheet.ExcludeFromCount = oExcludeSheets
		oSheet.ExcludeFromPrinting = oExcludeSheets
	End If
Next